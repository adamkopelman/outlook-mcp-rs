use std::sync::Arc;

use clap::Parser;
use outlook_mcp_rs::cli::{Cli, Mode};
use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;
use outlook_mcp_rs::transport::run_http;
use rmcp::{transport::stdio, ServiceExt};

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let mode = Cli::parse().resolve().map_err(|e| anyhow::anyhow!(e))?;

    let client = Arc::new(WindowsOutlookClient::new());
    let server = OutlookMcpServer::new(client);

    match mode {
        Mode::Stdio => {
            // Unchanged from the original: serve MCP over stdio.
            let service = server.serve(stdio()).await?;
            service.waiting().await?;
        }
        Mode::Http { addr, token } => {
            run_http(server, addr, token).await?;
        }
    }
    Ok(())
}
