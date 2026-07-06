use std::sync::Arc;

use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;
use rmcp::{transport::stdio, ServiceExt};

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let client = Arc::new(WindowsOutlookClient::new());
    let server = OutlookMcpServer::new(client);
    let service = server.serve(stdio()).await?;
    service.waiting().await?;
    Ok(())
}
