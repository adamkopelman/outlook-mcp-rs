//! End-to-end test of the Streamable HTTP transport against a fake Outlook
//! backend. Boots the real axum router on an ephemeral loopback port and
//! drives it with rmcp's streamable-HTTP client. No real Outlook required —
//! runs in CI like any other test.

use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::FakeOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;
use outlook_mcp_rs::transport::{build_router, MCP_PATH};
use rmcp::model::CallToolRequestParams;
use rmcp::transport::streamable_http_client::{
    StreamableHttpClientTransport, StreamableHttpClientTransportConfig,
};
use rmcp::ServiceExt;

/// Bind an ephemeral loopback port, serve `build_router` on it with the given
/// token, and return the base `http://127.0.0.1:PORT` URL. The server task
/// runs detached for the duration of the test process.
async fn spawn_server(token: Option<String>) -> String {
    let server = OutlookMcpServer::new(Arc::new(FakeOutlookClient::new()));
    let router = build_router(server, token);
    let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
    let addr = listener.local_addr().unwrap();
    tokio::spawn(async move {
        axum::serve(listener, router).await.unwrap();
    });
    format!("http://{addr}")
}

#[tokio::test]
async fn lists_and_calls_tools_over_http_with_correct_token() {
    let base = spawn_server(Some("s3cret".into())).await;
    let uri = format!("{base}{MCP_PATH}");

    let transport = StreamableHttpClientTransport::from_config(
        StreamableHttpClientTransportConfig::with_uri(uri).auth_header("s3cret"),
    );
    let client = ().serve(transport).await.expect("handshake should succeed with correct token");

    // The full tool surface is advertised.
    let tools = client.list_all_tools().await.expect("list_tools should succeed");
    assert!(
        tools.iter().any(|t| t.name == "list_folders"),
        "expected list_folders in advertised tools, got: {:?}",
        tools.iter().map(|t| &t.name).collect::<Vec<_>>()
    );

    // A tool call round-trips to the fake backend and back.
    let result = client
        .call_tool(CallToolRequestParams::new("list_folders"))
        .await
        .expect("call_tool(list_folders) should succeed");
    assert!(!result.content.is_empty(), "expected non-empty tool result");

    client.cancel().await.ok();
}

#[tokio::test]
async fn rejects_connection_without_token() {
    let base = spawn_server(Some("s3cret".into())).await;
    let uri = format!("{base}{MCP_PATH}");

    // No auth_header set → the server's middleware returns 401 before MCP,
    // so the handshake must fail.
    let transport = StreamableHttpClientTransport::from_uri(uri);
    let result = ().serve(transport).await;
    assert!(result.is_err(), "handshake must fail when the required token is absent");
}

// ---- UTF-8 on the wire, issue #31 ------------------------------------------

const HE_SUBJECT: &str = "מייל שיקוף";
const HE_SENDER: &str = "עדה לאבלייס";

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack.windows(needle.len()).position(|w| w == needle)
}

/// The complete chunks of a `Transfer-Encoding: chunked` body received so far.
fn dechunk(mut data: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    while let Some(eol) = find(data, b"\r\n") {
        let size_field = std::str::from_utf8(&data[..eol]).unwrap();
        let size = usize::from_str_radix(size_field.split(';').next().unwrap().trim(), 16).unwrap();
        let start = eol + 2;
        if size == 0 || data.len() < start + size + 2 {
            break;
        }
        out.extend_from_slice(&data[start..start + size]);
        data = &data[start + size + 2..];
    }
    out
}

/// POSTs `body` to the MCP endpoint over a raw socket, so the test sees the
/// exact response header and body bytes with no client-side decoding. Reads
/// until `done(body_so_far)` holds. Returns the header block and the
/// de-chunked body.
async fn raw_post(
    base: &str,
    session: Option<&str>,
    body: &serde_json::Value,
    done: impl Fn(&[u8]) -> bool,
) -> (String, Vec<u8>) {
    use tokio::io::{AsyncReadExt, AsyncWriteExt};

    let host = base.trim_start_matches("http://");
    let mut stream = tokio::net::TcpStream::connect(host).await.unwrap();
    let payload = serde_json::to_vec(body).unwrap();
    let mut request = format!(
        "POST {MCP_PATH} HTTP/1.1\r\nHost: {host}\r\nContent-Type: application/json\r\n\
         Accept: application/json, text/event-stream\r\nContent-Length: {}\r\n",
        payload.len()
    );
    if let Some(id) = session {
        request.push_str(&format!("Mcp-Session-Id: {id}\r\nMCP-Protocol-Version: 2025-06-18\r\n"));
    }
    request.push_str("\r\n");
    stream.write_all(request.as_bytes()).await.unwrap();
    stream.write_all(&payload).await.unwrap();

    let read = async {
        let mut raw = Vec::new();
        loop {
            let mut buf = [0u8; 4096];
            let n = stream.read(&mut buf).await.unwrap();
            raw.extend_from_slice(&buf[..n]);
            if let Some(end) = find(&raw, b"\r\n\r\n") {
                let head = String::from_utf8(raw[..end].to_vec()).expect("headers must be ASCII");
                let rest = &raw[end + 4..];
                let body = if head.to_ascii_lowercase().contains("transfer-encoding: chunked") {
                    dechunk(rest)
                } else {
                    rest.to_vec()
                };
                if n == 0 || done(&body) {
                    return (head, body);
                }
            }
            assert!(n > 0, "connection closed before the response was complete");
        }
    };
    tokio::time::timeout(std::time::Duration::from_secs(10), read)
        .await
        .expect("timed out waiting for the HTTP response")
}

fn header<'a>(head: &'a str, name: &str) -> Option<&'a str> {
    head.lines().find_map(|line| {
        let (k, v) = line.split_once(':')?;
        k.trim().eq_ignore_ascii_case(name).then(|| v.trim())
    })
}

/// A complete SSE event carrying the JSON-RPC response with `id`.
fn has_response(body: &[u8], id: u64) -> bool {
    let needle = format!("\"id\":{id}");
    find(body, needle.as_bytes()).is_some_and(|at| find(&body[at..], b"\n\n").is_some())
}

/// Issue #31: a client that falls back to latin1 for a `text/*` response
/// without a charset turns UTF-8 Hebrew into mojibake. Every MCP response
/// over HTTP must name `charset=utf-8`, and the bytes must be strict UTF-8
/// with the Hebrew written literally.
#[tokio::test]
async fn http_responses_declare_utf8_and_carry_raw_utf8_hebrew() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_text(HE_SUBJECT, HE_SENDER, "גוף");
    let router = build_router(OutlookMcpServer::new(fake.clone()), None);
    let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
    let base = format!("http://{}", listener.local_addr().unwrap());
    tokio::spawn(async move {
        axum::serve(listener, router).await.unwrap();
    });

    let init = serde_json::json!({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
        "protocolVersion": "2025-06-18", "capabilities": {},
        "clientInfo": {"name": "utf8-test", "version": "0"}}});
    let (head, _) = raw_post(&base, None, &init, |b| has_response(b, 1)).await;
    let content_type = header(&head, "content-type").expect("initialize response has a Content-Type");
    assert!(
        content_type.to_ascii_lowercase().contains("charset=utf-8"),
        "initialize Content-Type must name UTF-8: {content_type}"
    );
    let session = header(&head, "mcp-session-id").expect("stateful server returns a session id");

    let initialized = serde_json::json!({"jsonrpc": "2.0", "method": "notifications/initialized"});
    raw_post(&base, Some(session), &initialized, |_| true).await;

    let call = serde_json::json!({"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {
        "name": "list_emails", "arguments": {"query": HE_SUBJECT}}});
    let (head, body) = raw_post(&base, Some(session), &call, |b| has_response(b, 2)).await;
    let content_type = header(&head, "content-type").expect("tools/call response has a Content-Type");
    assert!(
        content_type.to_ascii_lowercase().contains("charset=utf-8"),
        "tools/call Content-Type must name UTF-8: {content_type}"
    );
    let text = String::from_utf8(body).expect("HTTP response body must be valid UTF-8");
    assert!(text.contains(HE_SUBJECT) && text.contains(HE_SENDER), "Hebrew not verbatim in {text}");
    assert!(!text.contains("\\u05"), "unexpected \\u escape in {text}");
    // Hebrew in the request reached the client unchanged too.
    assert_eq!(fake.calls()[0].1["query"], HE_SUBJECT);
}

/// Send one raw HTTP/1.1 request to `base` and return the response's status
/// line and body text (the server closes the connection after replying).
async fn raw_request(base: &str, head: &str, body: &[u8]) -> (String, String) {
    use tokio::io::{AsyncReadExt, AsyncWriteExt};
    let addr = base.trim_start_matches("http://");
    let mut stream = tokio::net::TcpStream::connect(addr).await.unwrap();
    stream.write_all(head.as_bytes()).await.unwrap();
    // The server may answer (and close) before the body is fully written.
    let _ = stream.write_all(body).await;
    let mut response = Vec::new();
    let _ = stream.read_to_end(&mut response).await;
    let text = String::from_utf8_lossy(&response).into_owned();
    let status = text.lines().next().unwrap_or("").to_string();
    let body = text.split_once("\r\n\r\n").map(|(_, b)| b.to_string()).unwrap_or_default();
    (status, body)
}

#[tokio::test]
async fn oversized_request_gets_an_explicit_payload_too_large() {
    use outlook_mcp_rs::transport::MAX_REQUEST_BODY_BYTES;
    let base = spawn_server(None).await;
    let head = format!(
        "POST {MCP_PATH} HTTP/1.1\r\nHost: localhost\r\nContent-Type: application/json\r\n\
         Accept: application/json, text/event-stream\r\nContent-Length: {}\r\nConnection: close\r\n\r\n",
        MAX_REQUEST_BODY_BYTES + 1
    );
    let (status, body) = raw_request(&base, &head, b"").await;
    assert!(status.contains("413"), "expected 413, got {status:?}");
    assert!(body.contains("payload too large") && body.contains("html_body_file"), "{body}");
}

#[tokio::test]
async fn compressed_request_is_refused_explicitly() {
    let base = spawn_server(None).await;
    let head = format!(
        "POST {MCP_PATH} HTTP/1.1\r\nHost: localhost\r\nContent-Type: application/json\r\n\
         Content-Encoding: gzip\r\nContent-Length: 4\r\nConnection: close\r\n\r\n"
    );
    let (status, body) = raw_request(&base, &head, b"\x1f\x8b\x08\x00").await;
    assert!(status.contains("415"), "expected 415, got {status:?}");
    assert!(body.contains("Content-Encoding"), "{body}");
}

/// Issue #28: a tool call carrying a large HTML body with embedded base64
/// images (76 KB in the report; ~1.5 MB here) goes through the HTTP transport.
#[tokio::test]
async fn large_html_body_tool_call_round_trips_over_http() {
    let base = spawn_server(None).await;
    let transport = StreamableHttpClientTransport::from_uri(format!("{base}{MCP_PATH}"));
    let client = ().serve(transport).await.expect("handshake");
    const PNG_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    let html = format!(
        "<p>{}</p><img src=\"data:image/png;base64,{PNG_B64}\">",
        "lorem ipsum ".repeat(130_000)
    );
    let args = serde_json::json!({"email_id": "entry-1|store-1", "html_body": html});
    let result = client
        .call_tool(CallToolRequestParams::new("update_draft").with_arguments(args.as_object().unwrap().clone()))
        .await
        .expect("a ~1.5 MB update_draft call should succeed");
    assert_ne!(result.is_error, Some(true), "{result:?}");
    client.cancel().await.ok();
}
