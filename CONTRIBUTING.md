# Contributing to outlook-mcp-rs

Thanks for your interest in contributing! This is a single-binary MCP server that drives
the **classic Microsoft Outlook desktop app** on Windows via its native COM automation API.
The notes below should get you productive quickly.

## Ways to contribute

- **Report a bug** — open a [bug report](https://github.com/adamkopelman/outlook-mcp-rs/issues/new?template=1-bug_report.yml).
- **Request a feature or new tool** — open a [feature request](https://github.com/adamkopelman/outlook-mcp-rs/issues/new?template=2-feature_request.yml).
- **Ask a question** — use [Discussions](https://github.com/adamkopelman/outlook-mcp-rs/discussions) rather than an issue.
- **Send a pull request** — see below.

Please search existing issues and pull requests before opening a new one to avoid duplicates.

## Development setup

### Requirements

- **Windows** — the server links against Outlook's Win32 COM API, so it only builds and runs
  meaningfully on Windows.
- **Classic Outlook desktop** installed, running, and signed in to a mailbox (the "new Outlook"
  is not supported — it does not expose the COM object model).
- A stable [Rust toolchain](https://rustup.rs/) (`rustup default stable`).

### Build

```sh
cargo build              # debug build
cargo build --release    # produces target/release/outlook-mcp-rs.exe
```

## Testing

The suite is deliberately split so that nothing is sent or delivered by accident. Tests whose
names begin with `live_outlook` talk to a real, running mailbox; everything else is safe to run
anywhere.

```sh
# Fast, safe: what CI runs. Skips the live-mailbox tests.
cargo test --all -- --skip live_outlook

# Full suite against a live Outlook mailbox (Windows, Outlook running & signed in).
cargo test --all
```

See [`TESTING.md`](TESTING.md) for details on the live tests and the safety model. Please run the
skip-live command before every PR; run the live suite too if your change touches COM interaction.

## Pull request process

1. **Branch** off `main` for your change; keep each PR focused on a single concern.
2. **Format and lint** before committing:
   ```sh
   cargo fmt
   cargo clippy --all -- -D warnings
   ```
3. **Test** with at least `cargo test --all -- --skip live_outlook`.
4. **Update docs** — if you add or change a tool, update the README tool list and any relevant docs.
5. **Open the PR** and fill in the pull request template, including the *Type of change*,
   *Side effects*, and *Testing* sections.

CI runs the test build on `windows-latest`; tagged pushes (`v*`) build and publish a release binary.

## Side effects: mail and mailbox writes

This project is deliberate about side effects. The handful of tools that actually **send mail or
meeting responses** — or otherwise **write to the mailbox** — must be explicit and opt-in, never a
silent consequence of a read. If your change adds or modifies a side-effecting tool:

- Keep the sending/writing action explicit and clearly named.
- Make sure tests do not deliver anything by accident (gate live behavior behind the
  `live_outlook` naming convention).
- Call it out in the PR's *Side effects* section.

## Coding conventions

- Keep code idiomatic Rust and consistent with the surrounding style; `cargo fmt` is the source of truth.
- Prefer clear, descriptive names for tools and their arguments — they are part of the public MCP surface.
- Never commit private mailbox content, credentials, or tokens in code, tests, fixtures, or logs.

## Code of conduct

Please be respectful and constructive in issues, discussions, and reviews. Assume good intent and
keep feedback focused on the work.

## License

By contributing, you agree that your contributions will be licensed under the same terms as the
project (see [`LICENSE`](LICENSE)).
