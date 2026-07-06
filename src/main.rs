use outlook_mcp_rs::server;

fn main() {
    println!("outlook-mcp-rs: scaffold only, real server wired in a later task");
    let _ = &server::OutlookMcpServer::new; // silence unused-import warning until Task 17
}
