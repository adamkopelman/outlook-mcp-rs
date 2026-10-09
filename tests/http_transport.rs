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
