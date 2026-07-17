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
