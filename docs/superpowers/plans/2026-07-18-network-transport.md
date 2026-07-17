# Optional Network (Streamable HTTP) Transport Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let the server optionally listen on the network over MCP Streamable HTTP (so a Claude client on another LAN machine can reach it at `hostname:port`), while keeping the existing local stdio mode as the unchanged zero-arg default.

**Architecture:** Add a `clap`-parsed CLI that resolves to a `Mode` (`Stdio` or `Http { addr, token }`). `main.rs` dispatches: stdio takes today's `serve(stdio())` path unchanged; http builds an `axum` `Router` that mounts rmcp's `StreamableHttpService` at `/mcp`, wrapped by a bearer-token auth middleware, served by `axum::serve`. No Outlook/COM/tool code changes — COM is already initialized per-call on a `spawn_blocking` thread, so concurrent HTTP requests are already safe. Full design: `docs/superpowers/specs/2026-07-18-network-transport-design.md`.

**Tech Stack:** Rust (2024 edition), `clap` 4 (derive), `axum` 0.8, `rmcp` 2.1.0 (`transport-streamable-http-server` feature), `constant_time_eq` 0.3, `tokio`. Tests use the existing `FakeOutlookClient` — no real Outlook, CI-safe.

## Global Constraints

- **stdio mode must stay byte-for-byte behaviorally unchanged**: `outlook-mcp-rs` with no arguments serves MCP over stdio exactly as today. Every existing stdio config keeps working.
- **Single self-contained binary**: new deps (`clap`, `axum`, `constant_time_eq`) are compile-time only; the shipped artifact stays one `.exe`.
- **No changes to any Outlook tool, the `OutlookClient` trait, COM code, or `OutlookMcpServer`'s tool surface.** This is a transport/entrypoint change only.
- **Network endpoint path is exactly `/mcp`.** Default bind when only `--port` given is `0.0.0.0:<port>`.
- **Auth is off unless `--token` is given**; when given, every request needs `Authorization: Bearer <token>`, compared in constant time; failures get `401` and never reach a tool.
- **`--token` is only valid in network mode**; `--port` and `--bind` are mutually exclusive; `--http` with neither `--port` nor `--bind` is an error.
- Match `axum` to rmcp's own version (**0.8**) to avoid `http`/`tower` version skew.

---

### Task 1: CLI parsing module (`src/cli.rs`)

Add `clap` and a pure, unit-testable `Cli`→`Mode` resolver. No I/O, no server — just argument semantics.

**Files:**
- Modify: `Cargo.toml` (add `clap`)
- Create: `src/cli.rs`
- Modify: `src/lib.rs` (add `pub mod cli;`)

**Interfaces:**
- Consumes: nothing from other tasks.
- Produces: `pub struct Cli` (clap `Parser`); `pub enum Mode { Stdio, Http { addr: std::net::SocketAddr, token: Option<String> } }`; `impl Cli { pub fn resolve(self) -> Result<Mode, String> }`. Task 3 calls `Cli::parse().resolve()`.

- [ ] **Step 1: Add the `clap` dependency**

In `Cargo.toml`, under `[dependencies]`, add (keep the existing entries; alphabetical placement near the top):

```toml
clap = { version = "4", features = ["derive"] }
```

- [ ] **Step 2: Write the failing tests**

Create `src/cli.rs` with ONLY the tests first (the types/`resolve` don't exist yet, so this won't compile — that's the RED state):

```rust
use std::net::SocketAddr;

use clap::Parser;

/// Command-line surface. No-arg invocation resolves to stdio (unchanged
/// default); `--http` (with `--port` or `--bind`) selects network mode.
#[derive(Parser, Debug)]
#[command(version, about = "MCP server for the classic Outlook desktop app (COM).")]
pub struct Cli {
    /// Run as a network server (Streamable HTTP) instead of stdio.
    #[arg(long)]
    pub http: bool,

    /// TCP port for --http; binds 0.0.0.0:<PORT>. Mutually exclusive with --bind.
    #[arg(long)]
    pub port: Option<u16>,

    /// Explicit bind socket address for --http, e.g. 0.0.0.0:8080.
    #[arg(long)]
    pub bind: Option<SocketAddr>,

    /// Require "Authorization: Bearer <TOKEN>" on every request (network mode only).
    #[arg(long)]
    pub token: Option<String>,
}

/// The resolved run mode after validating flag combinations.
#[derive(Debug, PartialEq, Eq)]
pub enum Mode {
    Stdio,
    Http { addr: SocketAddr, token: Option<String> },
}

impl Cli {
    /// Turn parsed flags into a validated `Mode`, or a human-readable error.
    /// Network mode is requested by `--http`, `--port`, or `--bind`.
    pub fn resolve(self) -> Result<Mode, String> {
        let network = self.http || self.port.is_some() || self.bind.is_some();
        if !network {
            if self.token.is_some() {
                return Err("--token only applies with --http".into());
            }
            return Ok(Mode::Stdio);
        }
        let addr = match (self.port, self.bind) {
            (Some(_), Some(_)) => {
                return Err("--port and --bind are mutually exclusive".into());
            }
            (Some(p), None) => SocketAddr::from(([0, 0, 0, 0], p)),
            (None, Some(a)) => a,
            (None, None) => return Err("--http requires --port or --bind".into()),
        };
        Ok(Mode::Http { addr, token: self.token })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cli(http: bool, port: Option<u16>, bind: Option<&str>, token: Option<&str>) -> Cli {
        Cli {
            http,
            port,
            bind: bind.map(|b| b.parse().unwrap()),
            token: token.map(String::from),
        }
    }

    #[test]
    fn no_args_is_stdio() {
        assert_eq!(cli(false, None, None, None).resolve().unwrap(), Mode::Stdio);
    }

    #[test]
    fn http_with_port_binds_all_interfaces() {
        let mode = cli(true, Some(8080), None, None).resolve().unwrap();
        assert_eq!(
            mode,
            Mode::Http { addr: "0.0.0.0:8080".parse().unwrap(), token: None }
        );
    }

    #[test]
    fn http_with_bind_uses_exact_addr() {
        let mode = cli(true, None, Some("127.0.0.1:9000"), None).resolve().unwrap();
        assert_eq!(
            mode,
            Mode::Http { addr: "127.0.0.1:9000".parse().unwrap(), token: None }
        );
    }

    #[test]
    fn port_or_bind_alone_implies_network_mode() {
        assert!(matches!(cli(false, Some(8080), None, None).resolve().unwrap(), Mode::Http { .. }));
        assert!(matches!(cli(false, None, Some("0.0.0.0:8080"), None).resolve().unwrap(), Mode::Http { .. }));
    }

    #[test]
    fn port_and_bind_together_is_error() {
        let err = cli(true, Some(8080), Some("0.0.0.0:9000"), None).resolve().unwrap_err();
        assert!(err.contains("mutually exclusive"));
    }

    #[test]
    fn http_without_port_or_bind_is_error() {
        let err = cli(true, None, None, None).resolve().unwrap_err();
        assert!(err.contains("requires --port or --bind"));
    }

    #[test]
    fn token_without_network_mode_is_error() {
        let err = cli(false, None, None, Some("secret")).resolve().unwrap_err();
        assert!(err.contains("--token only applies with --http"));
    }

    #[test]
    fn token_carried_into_http_mode() {
        let mode = cli(true, Some(8080), None, Some("secret")).resolve().unwrap();
        assert_eq!(
            mode,
            Mode::Http { addr: "0.0.0.0:8080".parse().unwrap(), token: Some("secret".into()) }
        );
    }
}
```

- [ ] **Step 3: Wire the module into the crate**

In `src/lib.rs`, add the module declaration (keep the others; alphabetical):

```rust
pub mod cli;
```

- [ ] **Step 4: Run the tests — they now compile and pass**

Because Step 2 already wrote both the implementation and the tests in one file (clap types can't be meaningfully stubbed separately), the "RED" here is a compile failure before `src/lib.rs` knows about the module and before `clap` is a dependency. Verify the transition:

Run: `cargo test --lib cli::`
Expected: the `cli::tests` module compiles and all 8 tests pass. (If `clap` was not added in Step 1, you instead get an unresolved-import error — add it and re-run.)

- [ ] **Step 5: Commit**

```bash
git add Cargo.toml Cargo.lock src/cli.rs src/lib.rs
git commit -m "Add clap CLI that resolves to stdio or http mode"
```

---

### Task 2: Network transport module (`src/transport.rs`)

Add `axum` + `constant_time_eq`, enable rmcp's server-side streamable-HTTP feature, and build the network wiring: a pure `is_authorized` (unit-tested here) plus `build_router`/`run_http` (compiled here, proven end-to-end in Task 4).

**Files:**
- Modify: `Cargo.toml` (add `axum`, `constant_time_eq`; extend the `rmcp` feature list)
- Create: `src/transport.rs`
- Modify: `src/lib.rs` (add `pub mod transport;`)

**Interfaces:**
- Consumes: `crate::server::OutlookMcpServer` (existing, `Clone`).
- Produces:
  - `pub fn is_authorized(configured: Option<&str>, header: Option<&str>) -> bool`
  - `pub fn build_router(server: OutlookMcpServer, token: Option<String>) -> axum::Router`
  - `pub async fn run_http(server: OutlookMcpServer, addr: std::net::SocketAddr, token: Option<String>) -> anyhow::Result<()>`
  - `pub const MCP_PATH: &str = "/mcp"` (used by Task 4's test to form the client URI)

- [ ] **Step 1: Add dependencies and enable the rmcp server feature**

In `Cargo.toml`, under `[dependencies]`: add `axum` and `constant_time_eq`, and add `"transport-streamable-http-server"` to the existing `rmcp` feature list. The `rmcp` line becomes:

```toml
rmcp = { version = "2.1.0", features = ["server", "transport-io", "transport-streamable-http-server"] }
```

Add these two new entries (alphabetical placement):

```toml
axum = "0.8"
constant_time_eq = "0.3"
```

- [ ] **Step 2: Write the failing auth tests**

Create `src/transport.rs` with the full module (implementation + tests). The RED state is that `src/lib.rs` does not yet declare the module and the new deps may be missing; the GREEN state is the `is_authorized` tests passing.

```rust
use std::net::SocketAddr;
use std::sync::Arc;

use axum::{
    body::Body,
    extract::State,
    http::{header::AUTHORIZATION, Request, StatusCode},
    middleware::{self, Next},
    response::Response,
    Router,
};
use rmcp::transport::streamable_http_server::{
    session::local::LocalSessionManager, StreamableHttpServerConfig, StreamableHttpService,
};

use crate::server::OutlookMcpServer;

/// The single HTTP path the MCP Streamable HTTP endpoint is mounted at.
pub const MCP_PATH: &str = "/mcp";

/// Pure authorization decision, split out so it is unit-testable without a
/// live socket. `configured` is the token the server was started with
/// (`None` = auth disabled). `header` is the raw incoming `Authorization`
/// header value, if any.
pub fn is_authorized(configured: Option<&str>, header: Option<&str>) -> bool {
    let Some(secret) = configured else {
        return true; // auth disabled
    };
    let Some(header) = header else {
        return false; // auth required but no header
    };
    let Some(presented) = header.strip_prefix("Bearer ") else {
        return false; // wrong scheme
    };
    // Constant-time compare to avoid a timing side channel on the secret.
    constant_time_eq::constant_time_eq(presented.as_bytes(), secret.as_bytes())
}

/// Axum middleware: gate every request on `is_authorized`, else 401.
async fn auth_middleware(
    State(token): State<Arc<Option<String>>>,
    req: Request<Body>,
    next: Next,
) -> Response {
    let header = req
        .headers()
        .get(AUTHORIZATION)
        .and_then(|v| v.to_str().ok());
    if is_authorized(token.as_deref(), header) {
        next.run(req).await
    } else {
        Response::builder()
            .status(StatusCode::UNAUTHORIZED)
            .body(Body::empty())
            .expect("static empty 401 response is always valid")
    }
}

/// Build the axum router: mount rmcp's Streamable HTTP MCP service at
/// `/mcp`, wrapped by the bearer-auth layer. Exposed (not just used by
/// `run_http`) so integration tests can drive it on an ephemeral port.
///
/// `allowed_hosts` is deliberately disabled: clients connect via the Windows
/// computer name (whose `Host` header we cannot predict), our clients are
/// native (not browsers vulnerable to DNS rebinding), and the bearer token is
/// the real access control. See the design doc's "DNS rebinding" section.
pub fn build_router(server: OutlookMcpServer, token: Option<String>) -> Router {
    let service = StreamableHttpService::new(
        move || Ok(server.clone()),
        Arc::new(LocalSessionManager::default()),
        StreamableHttpServerConfig::default().disable_allowed_hosts(),
    );
    Router::new()
        .route_service(MCP_PATH, service)
        .layer(middleware::from_fn_with_state(Arc::new(token), auth_middleware))
}

/// Bind `addr` and serve the MCP endpoint until the process is terminated.
pub async fn run_http(
    server: OutlookMcpServer,
    addr: SocketAddr,
    token: Option<String>,
) -> anyhow::Result<()> {
    let router = build_router(server, token);
    let listener = tokio::net::TcpListener::bind(addr).await?;
    // stderr, not stdout — never interferes with an MCP stdio stream (which
    // this mode isn't using anyway) and gives the operator a confirmation line.
    eprintln!("outlook-mcp-rs listening on http://{addr}{MCP_PATH}");
    axum::serve(listener, router).await?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::is_authorized;

    #[test]
    fn no_token_configured_allows_everything() {
        assert!(is_authorized(None, None));
        assert!(is_authorized(None, Some("Bearer whatever")));
        assert!(is_authorized(None, Some("garbage")));
    }

    #[test]
    fn correct_bearer_token_is_authorized() {
        assert!(is_authorized(Some("s3cret"), Some("Bearer s3cret")));
    }

    #[test]
    fn wrong_token_is_rejected() {
        assert!(!is_authorized(Some("s3cret"), Some("Bearer nope")));
    }

    #[test]
    fn missing_header_is_rejected_when_token_configured() {
        assert!(!is_authorized(Some("s3cret"), None));
    }

    #[test]
    fn wrong_scheme_is_rejected() {
        assert!(!is_authorized(Some("s3cret"), Some("Basic s3cret")));
        assert!(!is_authorized(Some("s3cret"), Some("s3cret")));
    }
}
```

- [ ] **Step 3: Wire the module into the crate**

In `src/lib.rs`, add (alphabetical, after `pub mod server;`):

```rust
pub mod transport;
```

- [ ] **Step 4: Run the auth unit tests**

Run: `cargo test --lib transport::`
Expected: the 5 `transport::tests` cases compile and pass. This also forces `build_router`/`run_http` to type-check against the real rmcp/axum APIs (a compile error here means a dependency/feature/signature mismatch to fix before moving on).

- [ ] **Step 5: Confirm the whole crate still builds**

Run: `cargo build`
Expected: clean build. (`run_http`/`build_router` have no dedicated unit test — they are proven end-to-end in Task 4.)

- [ ] **Step 6: Commit**

```bash
git add Cargo.toml Cargo.lock src/transport.rs src/lib.rs
git commit -m "Add streamable-HTTP transport with bearer-auth middleware"
```

---

### Task 3: Dispatch in `main.rs`

Wire the CLI to the two transports. This is the only change to the binary entrypoint.

**Files:**
- Modify: `src/main.rs`

**Interfaces:**
- Consumes: `outlook_mcp_rs::cli::{Cli, Mode}` (Task 1), `outlook_mcp_rs::transport::run_http` (Task 2), existing `OutlookMcpServer`/`WindowsOutlookClient`.
- Produces: nothing consumed by later tasks.

- [ ] **Step 1: Replace `main.rs` with the dispatcher**

Overwrite `src/main.rs` with:

```rust
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
```

- [ ] **Step 2: Verify stdio mode still starts (manual smoke, Windows)**

Run: `cargo run -- --help`
Expected: clap prints usage listing `--http`, `--port`, `--bind`, `--token`, `--help`, `--version`. Exit 0.

Run: `cargo run -- --http` (expect a clean validation error, not a panic)
Expected: stderr shows `--http requires --port or --bind`; non-zero exit.

Run: `cargo run -- --token x`
Expected: stderr shows `--token only applies with --http`; non-zero exit.

(Full stdio serving is exercised by the existing test suite and the live tests; no need to hand-drive an MCP stdio session here.)

- [ ] **Step 3: Run the whole unit suite**

Run: `cargo test`
Expected: all existing tests plus Task 1 & 2's new tests pass; no regressions.

- [ ] **Step 4: Commit**

```bash
git add src/main.rs
git commit -m "Dispatch main to stdio or http transport by CLI mode"
```

---

### Task 4: HTTP end-to-end integration test (`tests/http_transport.rs`)

Prove the whole network path — bind → axum → auth layer → rmcp service → `OutlookMcpServer` → `FakeOutlookClient` — with the rmcp streamable-HTTP *client*, including that auth actually gates the endpoint. No real Outlook.

**Files:**
- Modify: `Cargo.toml` (add a `[dev-dependencies]` rmcp entry enabling the client transport)
- Create: `tests/http_transport.rs`

**Interfaces:**
- Consumes: `build_router` + `MCP_PATH` (Task 2), `OutlookMcpServer` + `FakeOutlookClient` (existing, both public — already imported by `tests/tools.rs`).
- Produces: nothing.

- [ ] **Step 1: Add the client-transport dev-dependency**

In `Cargo.toml`, add a `[dev-dependencies]` section (or extend it if one exists) with rmcp's client features so the test can act as an MCP client. Cargo unions these with the normal `rmcp` features for test builds:

```toml
[dev-dependencies]
rmcp = { version = "2.1.0", features = ["client", "transport-streamable-http-client-reqwest"] }
```

- [ ] **Step 2: Write the integration test**

Create `tests/http_transport.rs`:

```rust
//! End-to-end test of the Streamable HTTP transport against a fake Outlook
//! backend. Boots the real axum router on an ephemeral loopback port and
//! drives it with rmcp's streamable-HTTP client. No real Outlook required —
//! runs in CI like any other test.

use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::FakeOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;
use outlook_mcp_rs::transport::{build_router, MCP_PATH};
use rmcp::model::CallToolRequestParam;
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
        .call_tool(CallToolRequestParam { name: "list_folders".into(), arguments: None })
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
```

- [ ] **Step 3: Run the integration test**

Run: `cargo test --test http_transport -- --nocapture`
Expected: both tests pass — `lists_and_calls_tools_over_http_with_correct_token` and `rejects_connection_without_token`. If `list_all_tools`/`call_tool`/`CallToolRequestParam` names differ in this rmcp version, adjust to the compiler's suggested paths (the client API is stable across 2.1.x; the types live under `rmcp::model` and the peer methods on the served client handle).

- [ ] **Step 4: Run the full suite to confirm no cross-test interference**

Run: `cargo test`
Expected: all tests green (unit + `tools` + `http_transport`), no port conflicts (each server binds `:0`).

- [ ] **Step 5: Commit**

```bash
git add Cargo.toml Cargo.lock tests/http_transport.rs
git commit -m "Add end-to-end streamable-HTTP transport test with auth gating"
```

---

### Task 5: Documentation and version bump

Document the new mode (both Windows and Ubuntu sides) and bump the version so the tag-driven release ships it.

**Files:**
- Modify: `README.md`
- Modify: `Cargo.toml` (version)

**Interfaces:**
- Consumes: the CLI/behavior finalized in Tasks 1–4.
- Produces: nothing.

- [ ] **Step 1: Add a "Remote / network mode" section to the README**

In `README.md`, immediately after the existing "Configure your MCP client" section (the stdio instructions, which stay as the primary/default path), insert:

```markdown
## Remote / network mode (connect from another machine)

By default the server speaks MCP over **stdio**, which requires the MCP client
to launch the binary as a local child process — so the client must run on the
same Windows machine as Outlook.

If your MCP client runs on a **different machine on the same LAN** (for example,
Claude on a Linux box talking to Outlook on your Windows PC), run the server in
**network mode** instead. It listens on a TCP port using MCP Streamable HTTP,
and the remote client connects by URL.

### On the Windows machine (where Outlook runs)

Start the server in HTTP mode with a shared secret token:

```
outlook-mcp-rs.exe --http --port 8080 --token YOUR_SECRET
```

It prints `outlook-mcp-rs listening on http://0.0.0.0:8080/mcp` and serves the
same 26 tools as stdio mode. Outlook must be running and signed in, as usual.

Find this machine's name (the client connects to it):

```
hostname
```

Allow the port through Windows Firewall, scoped to your local subnet (adjust
the port and subnet to match your network):

```
netsh advfirewall firewall add rule name="outlook-mcp-rs" dir=in action=allow ^
  protocol=TCP localport=8080 remoteip=192.168.1.0/24
```

### On the client machine (e.g. Ubuntu running Claude Code)

Point the client at `http://<WINDOWS-HOSTNAME>:8080/mcp`, sending the token as a
bearer header. With Claude Code:

```
claude mcp add --transport http outlook http://WINDOWS-PC:8080/mcp \
  --header "Authorization: Bearer YOUR_SECRET"
```

Replace `WINDOWS-PC` with the `hostname` from above (or the machine's LAN IP).
The Outlook tools then appear in that client.

### Options

| Flag | Meaning |
|---|---|
| `--http` | Enable network (Streamable HTTP) mode instead of stdio. |
| `--port <PORT>` | Listen on `0.0.0.0:<PORT>`. Mutually exclusive with `--bind`. |
| `--bind <ADDR>` | Listen on an exact socket address, e.g. `127.0.0.1:8080`. |
| `--token <SECRET>` | Require `Authorization: Bearer <SECRET>` on every request. |

### Security

- **Set a token.** Without `--token`, network mode accepts any request that
  reaches the port — anyone on the network could read or send your mail. The
  token is optional only to make first-run testing easy; use it for any real
  setup.
- **Scope the firewall** to the specific hosts/subnet that need access, as
  shown above.
- **Traffic is plain HTTP** (no TLS in this version). Keep it on a trusted LAN;
  if you need encryption across untrusted networks, front it with a reverse
  proxy or tunnel that terminates TLS.
```

- [ ] **Step 2: Update the intro line that claims stdio-only**

In `README.md`, the "Configure your MCP client" section opens with a sentence stating the server "speaks MCP over stdio, takes no arguments, and needs no environment variables." Change that sentence to acknowledge the new mode without disrupting the stdio-first framing. Replace:

```markdown
`outlook-mcp-rs` speaks MCP over stdio, takes no arguments, and needs no environment
variables. Point any MCP-capable client at the executable's path.
```

with:

```markdown
By default, `outlook-mcp-rs` speaks MCP over stdio and takes no arguments — point any
local MCP-capable client at the executable's path. (To connect from another machine on
your network, see [Remote / network mode](#remote--network-mode-connect-from-another-machine)
below.)
```

- [ ] **Step 3: Bump the version**

In `Cargo.toml`, change:

```toml
version = "0.2.0"
```

to:

```toml
version = "0.3.0"
```

- [ ] **Step 4: Confirm it all still builds and tests green**

Run: `cargo test`
Expected: full suite passes at the new version. (`cargo build` re-stamps `Cargo.lock` with `0.3.0`.)

- [ ] **Step 5: Commit**

```bash
git add README.md Cargo.toml Cargo.lock
git commit -m "Document network mode and bump version to 0.3.0"
```

---

## Self-Review Notes

- **Spec coverage:** stdio-unchanged default → Task 3 dispatch + Global Constraints; network Streamable HTTP mode → Tasks 2 & 4; CLI flags → Task 1; optional bearer token → Tasks 1 (carry), 2 (enforce), 4 (prove gating); `disable_allowed_hosts` rationale → Task 2 `build_router` doc; setup docs (Windows + Ubuntu) → Task 5; version bump/release → Task 5. Every spec section maps to a task.
- **Type consistency:** `Mode`/`Cli` defined in Task 1 are consumed verbatim in Task 3; `build_router`/`run_http`/`is_authorized`/`MCP_PATH` signatures defined in Task 2 are used unchanged in Tasks 3 and 4; `OutlookMcpServer::new(Arc<dyn OutlookClient>)` is the existing constructor used with `WindowsOutlookClient` (main) and `FakeOutlookClient` (test).
- **No placeholders:** every code and doc block is complete and literal; commands have expected output. The one intentionally flexible spot — exact rmcp client method names in Task 4 Step 3 — is called out with the fallback (follow the compiler) because it is the only surface not verifiable without compiling against the resolved crate.
- **Release note:** the release itself (tagging `v0.3.0` to trigger CI/CD) is deliberately left as a post-merge operator action, consistent with how `v0.2.0` shipped, not folded into an implementation task.
