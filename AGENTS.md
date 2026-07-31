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

`master` is protected: pushes go through a pull request, all six CI checks must
pass, and history stays linear. The repository allows **rebase merges only** —
squash and merge commits are disabled, so every commit reaches `master` intact.
Write commits worth keeping.

**Two lines are published.** `master` is 1.x, on the `latest` tag. The `0.2.x`
branch is 0.x, on the `legacy` tag, for callers who cannot take Node 20.19+ or a
`CompressionStream`-capable browser. A fix for output that is *wrong or
unopenable* usually belongs on both; anything that changes output which already
worked belongs on `master` alone.

`0.2.x` is **frozen but releasable**: nothing watches it, no dependency work
happens on it speculatively, and it still publishes a `0.2.z` when one is
warranted. So a release from it begins by catching the branch up — its
dependencies are unaudited and its action pins may have stopped working. See
[docs/RELEASING.md](docs/RELEASING.md).

## Conventions not enforced by tooling

Most of these are findings from an adversarial review pass, kept so the next pass
starts somewhere new. Two have run, both shipped a release, and none of what they
found was a crash — see
[docs/adversarial-review.md](docs/adversarial-review.md) before adding to this
list or starting a pass of your own.

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

- **Pin actions to a SHA and comment the *full* version, and re-read the comment
  on every bump.** `# v4.3.1`, never `# v4`. A major-only comment can survive a
  major bump and go on claiming the old major: a `checkout` pin read `# v4` while
  pointing at v7.0.1. The SHA is what runs, so this misleads a reader rather than
  the runner — which is why nothing catches it.

  Do not trust a rule about *when* Dependabot rewrites the comment. The obvious
  one — that it rewrites a spelled-out version and leaves a major-only one alone
  — held for `checkout` and `setup-node` and then failed: the two artifact bumps
  rewrote `# v4` to `# v7.0.1` and `# v8.0.1` on their own. Same repository, same
  week, same comment shape, opposite behaviour. Check the comment against the SHA
  when reviewing the bump; that is the only step that works either way.

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
- **[docs/adversarial-review.md](docs/adversarial-review.md)** — the pass that
  hunts for output that is plausible and wrong, what the two previous ones found,
  and the bug shapes that keep recurring. Read before a release that changes
  behaviour, or when a green suite feels like weaker evidence than it looks.
- **[docs/migrating-to-1.x.md](docs/migrating-to-1.x.md)** — the 0.2.x → 1.x
  migration, written to be followed mechanically. Read when a caller reports
  breakage after upgrading, or when changing anything it documents as stable.
- **[docs/RELEASING.md](docs/RELEASING.md)** — the release runbook. Read before
  changing the version or the publish workflow.
- **[SECURITY.md](SECURITY.md)** — threat model and reporting.
