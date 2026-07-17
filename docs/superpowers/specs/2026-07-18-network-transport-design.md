# Optional Network (Streamable HTTP) Transport — Design

**Date:** 2026-07-18
**Status:** Approved (brainstorming complete; ready for implementation planning)

## Problem

Some users run Claude (Claude Code / an MCP client) on a **Ubuntu machine** that
shares a LAN with the **Windows machine** where Outlook and this server live.
The server drives Outlook's desktop COM API, so it can only run on the Windows
box — it cannot run on the Ubuntu machine at all. Today the binary speaks MCP
only over **stdio**, which requires the client to launch the binary as a local
child process. A client on a different machine has no way to connect.

We need the server to *optionally* listen on the network so a remote client can
reach it at `hostname:port` (the Windows computer name + a port), while keeping
the existing local stdio mode as the unchanged default.

## Goals

- Keep **stdio mode** exactly as it is today (zero-arg invocation, unchanged
  behavior, existing configs keep working).
- Add a **network mode** that listens on a TCP port using MCP **Streamable
  HTTP**, reachable from another machine on the LAN via the Windows computer
  name.
- Support **optional bearer-token authentication** in network mode.
- Ship setup documentation for the new mode (Windows side: run + firewall;
  Ubuntu side: point Claude at the URL).
- One binary supports **both** modes; the mode is chosen at launch.

## Non-Goals

- No TLS/HTTPS in this iteration (LAN + bearer token + firewall is the trust
  model; TLS can be layered later behind a reverse proxy if wanted).
- No change to any Outlook tool, COM code, or the `OutlookClient` trait.
- No multi-user / per-user auth, no OAuth — a single shared bearer token only.
- No Windows service / auto-start integration (documented as a manual run).

## Decisions (from brainstorming)

| Decision | Choice | Rationale |
|---|---|---|
| Network transport | **Streamable HTTP** | Current MCP standard; what `claude mcp add --transport http` uses; SSE is being phased out of the spec. |
| Mode selection | **CLI flags** (via `clap`) | Explicit, discoverable via `--help`, easy to document. No-arg stays stdio. |
| Auth | **Optional bearer token** | Off by default for easy first setup, strongly recommended in docs; stdio never needs it. |
| Arg parsing | **Add `clap`** (derive) | Idiomatic, free `--help`/validation, statically linked (single-binary story intact). |
| Default bind | **`0.0.0.0`** when only `--port` given | Remote reachability is the whole point; pair with token + firewall. |
| `Host` validation | **Disabled in network mode** | See "DNS rebinding" below. |

## CLI Surface

```
outlook-mcp-rs [OPTIONS]

Options:
      --http            Run as a network server (Streamable HTTP) instead of stdio.
      --port <PORT>     TCP port for --http. Binds 0.0.0.0:<PORT>.
      --bind <ADDR>     Explicit bind socket address for --http (e.g. 0.0.0.0:8080).
                        Mutually exclusive with --port.
      --token <TOKEN>   Require "Authorization: Bearer <TOKEN>" on every request
                        (network mode only). If omitted, no auth is enforced.
  -h, --help
  -V, --version
```

Behavior / validation rules:

- **No args → stdio mode** (today's default, unchanged).
- `--http` selects network mode. `--port` or `--bind` imply/require network
  mode.
- `--port` and `--bind` are **mutually exclusive**. `--http --port 8080` →
  bind `0.0.0.0:8080`. `--http --bind 127.0.0.1:8080` → bind exactly that.
- `--http` with neither `--port` nor `--bind` → error (no implicit default
  port; force an explicit choice so nothing silently opens a surprise port).
- `--token` outside network mode (stdio) → error ("--token only applies with
  --http"); it is meaningless over stdio.
- `--port`/`--bind`/`--token` given without `--http` → error, telling the user
  to add `--http`.

Example launches:

```
outlook-mcp-rs                                   # stdio (local, as today)
outlook-mcp-rs --http --port 8080                # LAN, no auth (warned in docs)
outlook-mcp-rs --http --port 8080 --token s3cret # LAN, bearer auth (recommended)
outlook-mcp-rs --http --bind 127.0.0.1:9000      # HTTP but loopback-only
```

## Architecture

The change is confined to the transport/entrypoint layer. The MCP server
(`OutlookMcpServer`), every tool, and all COM code are untouched.

```
                         ┌─────────────────────────────────────────┐
 args ── clap ──► Cli ──►│  main.rs dispatch                        │
                         │   ├── stdio  → server.serve(stdio())     │  (existing path)
                         │   └── http   → run_http(server, opts)    │  (new)
                         └─────────────────────────────────────────┘
                                              │
                     run_http builds an axum Router:
                       Router
                         .route_service("/mcp", StreamableHttpService::new(
                              factory: || Ok(server.clone()),
                              LocalSessionManager::default(),
                              StreamableHttpServerConfig::default()
                                  .disable_allowed_hosts() ))
                         .layer(bearer_auth_layer(token))     ← new middleware
                     served by axum::serve(TcpListener::bind(addr), router)
```

### Components / file structure

- **`src/cli.rs` (new)** — a `clap`-derived `Cli` struct plus a `Mode` it
  resolves to. One responsibility: turn `argv` into a validated
  `Mode::Stdio` or `Mode::Http { addr: SocketAddr, token: Option<String> }`,
  or a clear error. Pure and unit-testable (no I/O). Interface:
  `Cli::parse()` (clap) → `cli.resolve() -> Result<Mode, String>`.
- **`src/transport.rs` (new)** — network-mode wiring. One responsibility:
  given the built `OutlookMcpServer`, a `SocketAddr`, and an optional token,
  build the axum router (mount `StreamableHttpService` at `/mcp`, attach the
  auth layer) and serve it. Interface:
  `async fn run_http(server: OutlookMcpServer, addr: SocketAddr, token: Option<String>) -> anyhow::Result<()>`.
  Also holds the auth check as a pure function
  `fn is_authorized(configured: Option<&str>, header: Option<&str>) -> bool`
  so the token logic is unit-testable without a live socket.
- **`src/main.rs` (modified)** — thin dispatcher: `Cli::parse()` →
  `cli.resolve()` → build `OutlookMcpServer` (as today) → match `Mode` to
  either the existing `serve(stdio())` path or `transport::run_http(...)`.
- **`src/lib.rs` (modified)** — declare `pub mod cli;` and
  `pub mod transport;` so integration tests can drive them.
- **`Cargo.toml` (modified)** — add `clap` (derive), `axum` (0.8, matching
  rmcp), and enable rmcp's `transport-streamable-http-server` feature; add
  dev-only client feature for the integration test (see Testing).

### Authentication

A single axum middleware layer (`middleware::from_fn_with_state` or a small
`tower` layer) wraps the `/mcp` service:

- If **no token** is configured → pass every request through unchanged.
- If a **token** is configured → read the `Authorization` header; the request
  passes only when it equals `Bearer <token>`, compared in **constant time**
  (avoid a timing side channel on the secret). Otherwise respond **401
  Unauthorized** with an empty body and never touch the MCP service.

The comparison logic lives in the pure `is_authorized(configured, header)`
function so it is tested directly (match / mismatch / missing header / no token
configured), independent of axum.

### DNS rebinding / `allowed_hosts` (why it's disabled)

rmcp's `StreamableHttpServerConfig.allowed_hosts` defaults to
`["localhost","127.0.0.1","::1"]` and rejects any request whose HTTP `Host`
header isn't in that list. This is a **defense-in-depth guard against
browser-based DNS-rebinding attacks** on *locally-running, unauthenticated*
servers: malicious web-page JavaScript that rebinds a domain to `127.0.0.1`
still sends `Host: evil.com`, which the check rejects.

It is the wrong tool for this feature, for two reasons:

1. **Our clients are native processes, not browsers.** DNS rebinding is a
   browser-sandbox-escape attack; a native MCP client that intends to reach
   `WINDOWS-PC:8080` is not the threat this guard defends against.
2. **We have real authentication instead.** Rebinding is dangerous precisely
   because the target has *no auth*. Network mode uses a bearer token, which is
   actual authentication rather than a "is this local?" heuristic.

Keeping the check on would also be brittle: a remote client connects using the
Windows computer name (or `.local` mDNS name, or a possibly-DHCP-changed IP),
and we cannot predict, at startup, every `Host` value a client might send —
each mismatch would be a cryptic hard rejection. So network mode calls
`.disable_allowed_hosts()`, and the real controls become: **bearer token**
(primary), **a LAN-scoped firewall rule** (documented), and the **trusted
network** assumption.

**Accepted trade-off:** running network mode *without* `--token` is a genuinely
open door (anyone who can route a packet to `host:port` gets full mailbox
control, with no `Host` backstop). This is why the token is "off by default but
strongly recommended," and the setup docs flag tokenless network mode as
LAN-only-at-your-own-risk.

### Concurrency / COM

No changes required. Each tool call already runs inside
`run_blocking` (`spawn_blocking`), where a `ComGuard` initializes an
apartment-threaded COM context on that specific OS thread for the duration of
the call and uninitializes on drop (`src/outlook/com.rs`, `src/server.rs`).
Concurrent HTTP requests therefore each get an independent COM apartment on an
independent blocking thread — the same isolation stdio already relies on. The
HTTP server adds concurrency the tool layer already tolerates.

## Error Handling

- **CLI validation errors** (mutually-exclusive flags, `--token` without
  `--http`, `--http` without a port/bind) → `clap`/`resolve()` produce a clear
  message to stderr and a non-zero exit, before any server starts.
- **Bind failure** (port in use, permission) → `run_http` returns the
  `std::io::Error` up through `anyhow` to `main`; the process exits non-zero
  with the OS message.
- **Auth failure** → `401 Unauthorized`, request never reaches a tool.
- **Tool/COM errors** → unchanged; they surface as MCP tool errors exactly as
  they do over stdio (same `OutlookMcpServer`).

## Testing Strategy

All automated tests use the existing `FakeOutlookClient` — no real Outlook, so
they run in CI on any platform.

1. **CLI resolution (unit, `src/cli.rs`).** Table-driven over `resolve()`:
   no args → `Stdio`; `--http --port 8080` → `Http` bound `0.0.0.0:8080`, no
   token; `--http --bind` → exact addr; `--port` + `--bind` together → error;
   `--http` with neither → error; `--token` without `--http` → error;
   `--token` with `--http` → token present.
2. **Auth logic (unit, `src/transport.rs`).** `is_authorized`: no token
   configured → always true; correct `Bearer` → true; wrong token → false;
   missing header → false; malformed header → false.
3. **HTTP end-to-end (integration, `tests/http_transport.rs`).** Boot
   `run_http` on `127.0.0.1:0` (ephemeral port) with a `FakeOutlookClient`
   server and a configured token. Using rmcp's **streamable-HTTP client**
   transport (dev-dependency feature, with `auth_header` set):
   - with the correct token → `initialize` + `list_tools` succeeds and returns
     the expected tool set; calling one tool (e.g. `list_folders`) returns the
     fake's canned result.
   - with a wrong/absent token → the connection is rejected (401), asserting
     auth actually gates the endpoint.
   This proves the whole network path (bind → axum → auth layer → rmcp service
   → `OutlookMcpServer` → fake client) without real Outlook.

Existing stdio behavior is already covered; a smoke check that the stdio path
still builds/serves is implicitly covered by the unchanged `main` dispatch plus
the existing suite.

## Documentation

`README.md` gains a "Remote / network mode" section, and/or a short companion
doc, covering:

- **Windows (server) side:** run `outlook-mcp-rs.exe --http --port 8080 --token <secret>`;
  find the machine name (`hostname`); add an inbound Windows Firewall rule for
  the port, scoped to the LAN subnet (example `netsh advfirewall` command);
  note that Outlook must be running and signed in as usual.
- **Ubuntu (client) side:** point Claude Code at the URL, e.g.
  `claude mcp add --transport http outlook http://WINDOWS-PC:8080/mcp --header "Authorization: Bearer <secret>"`
  (exact flag syntax verified during implementation), and confirm the tools
  appear.
- **Security note:** reiterate the token recommendation and the tokenless
  trade-off; state that traffic is unencrypted HTTP over the LAN (no TLS this
  iteration).
- Keep the existing stdio instructions as the default/primary path.

## Rollout / Compatibility

- Fully backward compatible: no-arg invocation and every existing stdio config
  are unchanged.
- New crate dependencies (`clap`, `axum`) are compile-time only; the shipped
  artifact is still a single self-contained `.exe`.
- Warrants a **minor** version bump (new feature, no breaking change), e.g.
  `0.3.0`, released via the existing tag-driven CI/CD.
