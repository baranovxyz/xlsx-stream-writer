# Releasing xlsx-stream-writer

Maintainer runbook for publishing to npm. The pipeline is
`.github/workflows/publish.yml`; this file explains how to drive it and what has
to be true before it will work.

## Invariants

- **No local npm auth, ever.** Publishing uses npm OIDC trusted publishing from
  GitHub Actions. You never run `npm publish` from a laptop, and no long-lived
  npm token exists.
- **The workflow never bumps the version.** It publishes whatever version is in
  `package.json` on `master`, then tags and releases that commit. Bumping is a
  reviewed commit like any other.
- **The credential-holding job never sees repository code.** `prepare` installs,
  tests and packs; `publish` only downloads that tarball, checks its SHA-512
  against the value `prepare` recorded, and publishes it.
- **Every version is published once.** `prepare` fails if the version already
  exists on the registry.

## One-time setup

1. **Decide on repository visibility.** The npm CLI's own guard is about the
   *package*: it refuses `--provenance` unless the package is public or
   `--access public` is passed, which this workflow does
   (`libnpmpublish/lib/publish.js`: "Can't generate provenance for new or
   private package, you must set `access` to public").

   The repository is a separate question. `xlsx-stream-writer` is private on
   GitHub today, and a provenance attestation publicly records the repository
   URL, the workflow path and the release commit SHA in a transparency log. So
   publishing with provenance from a private repo advertises that the repo
   exists and which commit shipped. Make the repository public before the first
   provenance-backed release, or drop `--provenance` from `publish.yml` and
   accept an unattested release.

2. **Configure the trusted publisher on npmjs.org.** Package
   `xlsx-stream-writer` → Settings → Trusted publisher → GitHub Actions:

   | Field       | Value                                |
   | ----------- | ------------------------------------ |
   | Repository  | `baranovxyz/xlsx-stream-writer`      |
   | Workflow    | `.github/workflows/publish.yml`      |
   | Environment | `npm`                                |

3. **Create the `npm` environment** in the GitHub repository settings. Add
   required reviewers if you want a human gate before the token is issued.

4. **Remove any `NPM_TOKEN` secret.** Trusted publishing makes it unnecessary,
   and a lingering token is a credential to steal.

## Cutting a release

1. On a branch off `master`, in one commit:
   - bump `version` in `package.json`
   - add a dated section to `CHANGELOG.md`
2. `npm ci && npm test && npm audit --audit-level=low` — all clean.
3. `npm pack --dry-run` — confirm the file list is `dist/`, `src/`, `README.md`,
   `CHANGELOG.md`, `LICENSE`, `package.json` and nothing else. `src/` ships
   because `dist/` carries source maps that point at it.
4. Open a PR, get CI green, merge to `master`.
5. Dry run first:
   `gh workflow run publish.yml --repo baranovxyz/xlsx-stream-writer -f dry-run=true`
   A dry run exercises the whole preparation graph — install, test, audit, pack,
   file-list check, isolated install smoke test, secret scan — without touching
   the registry.
6. Publish:
   `gh workflow run publish.yml --repo baranovxyz/xlsx-stream-writer`
7. The graph must go green end to end: **prepare** → **publish** → **verify**
   (registry integrity plus SLSA provenance matching this workflow, this branch
   and this commit) → **tag-and-release**.

## After the release

- `npm view xlsx-stream-writer version` shows the new version.
- `npm view xlsx-stream-writer dist.attestations` shows a provenance URL.
- The `v<x.y.z>` tag and GitHub release exist and point at the release commit.

## When something goes wrong

- **`prepare` fails on the version guard.** The version is already on npm. Bump
  and start again; npm versions are immutable and must never be reused.
- **`publish` fails with a provenance error.** Either `--access public` is
  missing, or the trusted publisher entry does not match the repository,
  workflow path and environment exactly.
- **`publish` succeeded but `verify` failed.** The package is live — npm cannot
  be rolled back after 72 hours, and unpublishing a version burns it forever.
  Investigate before republishing; in most cases the fix is a new patch version,
  not an unpublish.
- **`tag-and-release` failed.** Cosmetic. Re-run the workflow, or create the tag
  and release by hand; the published artifact is unaffected.

## Release order for the hardening work

Each phase is independently releasable, and each is a real user-visible step:

| Phase | Version | Content                                                |
| ----- | ------- | ------------------------------------------------------ |
| 0 + 1 | 0.2.7   | Advisories cleared, test harness, CI, release pipeline |
| 2     | 0.3.0   | Correctness and robustness fixes                       |
| 3     | 1.0.0   | Built-in ZIP writer, zero runtime dependencies         |
| 5     | 1.1.0   | TypeScript sources and shipped type declarations       |

Phase 4 (supply-chain hardening) landed early, alongside 0.2.7, so that every
later phase publishes through the verified pipeline rather than by hand.

All five are committed. Publishing any of them needs the one-time setup above:
the trusted publisher entry on npmjs.org and the `npm` environment on GitHub.
Neither exists yet, so no release has been published.

If you would rather publish only the end state, release 1.1.0 and skip the
intermediate versions — the changelog documents each step regardless. If you
want the full history on npm, publish them in order, bumping `package.json` to
each version on `master` in turn.
