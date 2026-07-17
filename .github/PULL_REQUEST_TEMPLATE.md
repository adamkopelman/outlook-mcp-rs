<!--
Thanks for contributing to outlook-mcp-rs! Please fill out the sections below.
Keep the PR focused — one logical change per PR is easier to review.
-->

## Summary

<!-- What does this PR do, and why? Link any related issue, e.g. "Closes #123". -->

## Type of change

<!-- Put an "x" in the boxes that apply. -->

- [ ] 🐞 Bug fix (non-breaking change that fixes an issue)
- [ ] ✨ New feature / tool (non-breaking change that adds functionality)
- [ ] 💥 Breaking change (fix or feature that changes existing tool behavior or arguments)
- [ ] 📖 Documentation only
- [ ] 🧹 Refactor / internal (no user-facing change)
- [ ] 🔧 CI / build / tooling

## Changes

<!-- Bullet list of the notable changes. Mention any new or modified MCP tools and their arguments. -->

-

## Side effects

<!-- Does this PR add or change any tool that sends mail, responds to meetings, or writes to the mailbox? -->

- [ ] This PR does **not** introduce or change any sending / mailbox-writing behavior
- [ ] This PR **does** change a side-effecting tool (explain below and confirm it stays explicit / opt-in)

## Testing

<!-- How did you verify this? Note whether live Outlook tests were run. -->

- [ ] `cargo test --all -- --skip live_outlook` passes locally
- [ ] Verified against a live Outlook mailbox (describe below)
- [ ] `cargo fmt` / `cargo clippy` are clean

<!-- Describe the manual/live testing you performed, if any. -->

## Checklist

- [ ] I have read the contributing guidance and this change is focused on a single concern.
- [ ] Documentation (README / tool list) is updated if behavior or tools changed.
- [ ] No private mailbox content or credentials are included in the diff, tests, or logs.
