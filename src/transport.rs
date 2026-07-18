use std::net::SocketAddr;
use std::sync::Arc;

use axum::{
    body::Body,
    extract::State,
    http::{
        header::{AUTHORIZATION, WWW_AUTHENTICATE},
        Request, StatusCode,
    },
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
    // Constant-time compare to avoid a timing side channel on the secret's
    // *content*. `constant_time_eq` still short-circuits on a length
    // mismatch, so token *length* is technically observable via timing —
    // negligible against a reasonably long, high-entropy token.
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
            .header(WWW_AUTHENTICATE, "Bearer")
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
