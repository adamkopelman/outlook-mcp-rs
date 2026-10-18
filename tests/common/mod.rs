//! Shared harness for the system tests that run through the real MCP tool
//! layer: `OutlookMcpServer` over a real `WindowsOutlookClient`, served
//! in-process and called with JSON-RPC `tools/call` over rmcp's stdio codec,
//! plus the non-panicking results log and small JSON helpers.
#![allow(dead_code)]

use std::sync::Arc;

use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;
use serde_json::{json, Value};

/// Starts the real server in-process and returns a connected client.
pub async fn connect(name: &str) -> Wire {
    Wire::start(OutlookMcpServer::new(Arc::new(WindowsOutlookClient::new())), name).await
}

// ---- JSON-RPC over the stdio codec ------------------------------------------

pub struct Wire {
    reader: tokio::io::BufReader<tokio::io::ReadHalf<tokio::io::DuplexStream>>,
    writer: tokio::io::WriteHalf<tokio::io::DuplexStream>,
    next_id: u64,
}

impl Wire {
    pub async fn start(server: OutlookMcpServer, name: &str) -> Self {
        use rmcp::ServiceExt;
        let (client_io, server_io) = tokio::io::duplex(16 * 1024 * 1024);
        tokio::spawn(async move {
            if let Ok(running) = server.serve(server_io).await {
                let _ = running.waiting().await;
            }
        });
        let (read_half, writer) = tokio::io::split(client_io);
        let mut w = Wire { reader: tokio::io::BufReader::new(read_half), writer, next_id: 1 };
        w.request("initialize", json!({
            "protocolVersion": "2025-06-18", "capabilities": {},
            "clientInfo": {"name": name, "version": "0"}}))
            .await
            .expect("initialize");
        w.send(&json!({"jsonrpc": "2.0", "method": "notifications/initialized"})).await;
        w
    }

    pub async fn send(&mut self, msg: &Value) {
        use tokio::io::AsyncWriteExt;
        let mut bytes = serde_json::to_vec(msg).unwrap();
        bytes.push(b'\n');
        self.writer.write_all(&bytes).await.unwrap();
    }

    /// One request; Err if the wire text isn't UTF-8 or the server closed.
    pub async fn request(&mut self, method: &str, params: Value) -> Result<Value, String> {
        use tokio::io::AsyncBufReadExt;
        let id = self.next_id;
        self.next_id += 1;
        self.send(&json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params})).await;
        loop {
            let mut line = Vec::new();
            let n = self.reader.read_until(b'\n', &mut line).await.map_err(|e| e.to_string())?;
            if n == 0 {
                return Err("server closed".into());
            }
            let wire = String::from_utf8(line).map_err(|_| "response is not valid UTF-8".to_string())?;
            let msg: Value = serde_json::from_str(&wire).map_err(|e| e.to_string())?;
            if msg["id"] == id {
                return Ok(msg);
            }
        }
    }

    /// `tools/call`: Ok(parsed result) or Err(error text), whether the tool
    /// failed with a JSON-RPC error or an `isError` result.
    pub async fn call(&mut self, tool: &str, args: Value) -> Result<Value, String> {
        let msg = self.request("tools/call", json!({"name": tool, "arguments": args})).await?;
        if let Some(e) = msg.get("error") {
            return Err(e["message"].as_str().unwrap_or("").to_string());
        }
        let r = &msg["result"];
        let text = r["content"][0]["text"].as_str().unwrap_or("").to_string();
        if r["isError"] == true {
            return Err(text);
        }
        Ok(serde_json::from_str(&text).unwrap_or(Value::String(text)))
    }
}

// ---- results log ------------------------------------------------------------

#[derive(Default)]
pub struct Log(pub Vec<(String, bool, String)>);

impl Log {
    pub fn check(&mut self, id: &str, ok: bool, detail: impl Into<String>) -> bool {
        let detail = detail.into();
        println!("[{}] {id}: {detail}", if ok { "PASS" } else { "FAIL" });
        self.0.push((id.to_string(), ok, detail));
        ok
    }
    pub fn fail(&mut self, id: &str, detail: impl Into<String>) {
        self.check(id, false, detail);
    }
}

pub fn s(v: &Value, k: &str) -> String {
    v[k].as_str().unwrap_or("").to_string()
}

pub fn ids_of(rows: &Value) -> Vec<String> {
    rows.as_array().map(|a| a.iter().map(|r| s(r, "id")).collect()).unwrap_or_default()
}

pub fn subjects_of(rows: &Value) -> Vec<String> {
    rows.as_array().map(|a| a.iter().map(|r| s(r, "subject")).collect()).unwrap_or_default()
}

pub fn temp_dir(name: &str) -> std::path::PathBuf {
    let dir = std::env::temp_dir().join(format!("outlook-mcp-rs-systest-1017-{}-{name}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

/// Prints the summary table; true when every step passed.
impl Log {
    pub fn summary(&self) -> bool {
        let failed = self.0.iter().filter(|(_, ok, _)| !ok).count();
        for (id, ok, detail) in &self.0 {
            println!("{:<8} {} {}", id, if *ok { "PASS" } else { "FAIL" }, if *ok { "" } else { detail });
        }
        println!("{} / {} passed", self.0.len() - failed, self.0.len());
        failed == 0
    }
}
