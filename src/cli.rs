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
