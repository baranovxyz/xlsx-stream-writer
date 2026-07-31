# AGENTS.md

`xlsx-stream-writer` writes `.xlsx` files in streaming mode, in Node.js and the
browser. It is published to npm.

**Zero runtime dependencies is a feature, not an accident.** The ZIP container is
written here rather than delegated. Adding a runtime dependency is a deliberate
decision, not a convenience.

## Commands

| Task | Command |
| --- | --- |
| Test | `npm test` |
| Build only | `npm run build` |
| Type-check a consumer against the emitted declarations | `npm run test:types` |
| Audit | `npm audit --audit-level=low` |
| Preview the published tarball | `npm pack --dry-run` |

`npm test` compiles before it runs, because **the suite imports the built output,
not the sources**. Invoking the test runner directly exercises whatever was built
last, so a source edit followed by a bare runner invocation silently reports
stale results. Always go through `npm test`.

`master` is protected: pushes go through a pull request, all five CI checks must
pass, and history stays linear. The repository allows **rebase merges only** —
squash and merge commits are disabled, so every commit reaches `master` intact.
Write commits worth keeping.

## Conventions not enforced by tooling

- **Never key a lookup table by a caller-supplied string using a plain object.**
  Use a `Map`. Strings that name a member of `Object.prototype` resolve to the
  inherited member instead of missing, and that value then gets written into the
  document where an index belonged. This has produced corrupt workbooks twice,
  in two unrelated tables.

- **Never alias module-level mutable state into per-instance state.** Copy it.
  Sharing a module-level array across instances made the second workbook built in
  a process inherit the first one's styles, silently shifting every style index.

- **Anything caller-supplied that reaches XML must be escaped or validated
  first** — cell values, style attributes, and the numbers a caller's callback
  returns. An unvalidated value can close its attribute and open another.

- **No literal control characters in source files.** Write `\uXXXX` escapes, even
  inside test fixtures. Literal ones make tools treat the file as binary.

- **The golden fixtures are byte-exact and deliberate.** A diff against them is a
  behaviour change that belongs in the changelog. It is never a fixture to
  refresh so the suite goes green.

- **Excel reports "unreadable content", never the real cause.** A malformed part
  costs far more to diagnose than to prevent, so prefer failing loudly at the
  offending row over writing something questionable.

## Where to read more

- **[docs/architecture.md](docs/architecture.md)** — the design decisions the
  code does not explain on its own. Read before changing the ZIP writer, the
  stream plumbing, or the order parts are written in.
- **[docs/verifying-output.md](docs/verifying-output.md)** — how output
  correctness is actually established, and what remains unverified. Read before
  treating a green suite as proof that a workbook opens.
- **[docs/migrating-to-1.x.md](docs/migrating-to-1.x.md)** — the 0.2.x → 1.x
  migration, written to be followed mechanically. Read when a caller reports
  breakage after upgrading, or when changing anything it documents as stable.
- **[docs/RELEASING.md](docs/RELEASING.md)** — the release runbook. Read before
  changing the version or the publish workflow.
- **[SECURITY.md](SECURITY.md)** — threat model and reporting.
