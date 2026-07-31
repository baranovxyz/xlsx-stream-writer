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

## Changing the pipeline itself

**CI does not run this workflow.** A pull request that edits `publish.yml` shows
the same six green checks as any other, and not one of them executed a line of
it — they all exercise `ci.yml`. Treat that green as silence, not evidence.

Rehearsing an edit is harder than it looks, because two guards block the obvious
approaches:

- `prepare` runs only on `master` or a `*.x` branch, so dispatching the workflow
  against a feature branch skips the entire graph.
- The republish guard fails whenever the version in `package.json` is already on
  the registry. It reads `dry-run` into the environment but does not branch on
  it, so a dry run at the current version stops there too.

Together those mean an edit here cannot be proven before it reaches `master`,
and cannot be rehearsed on `master` at an already-published version. The first
honest test is the dry run of the next real release. So land pipeline changes at
the *start* of a release cycle, while there is still a version bump ahead of them
to exercise the change — never as the last commit before publishing.

## One-time setup

Already done for this package. This is here for a fresh fork or a rebuilt
repository — do not redo it against the live package.

1. **Configure the trusted publisher on npmjs.org.** Package
   `xlsx-stream-writer` → Settings → Trusted publisher → GitHub Actions:

   | Field       | Value                                |
   | ----------- | ------------------------------------ |
   | Repository  | `baranovxyz/xlsx-stream-writer`      |
   | Workflow    | `.github/workflows/publish.yml`      |
   | Environment | `npm`                                |

2. **Create the `npm` environment** in the GitHub repository settings. Add
   required reviewers if you want a human gate before the token is issued.

3. **Remove any `NPM_TOKEN` secret.** Trusted publishing makes it unnecessary,
   and a lingering token is a credential to steal.

## The 0.2 maintenance line

`^0.2.6` resolves to `>=0.2.6 <0.3.0`, so nothing on the 1.x line ever reaches a
project pinned to the old range. The `0.2.x` branch exists to serve those
projects, and a release from it must be a `0.2.z` patch — a `0.3.0` would not
reach them either.

Only backport fixes for input that was already producing a wrong or unopenable
file. Anything that changes the output of input that previously worked belongs
to 1.x, however much more correct it is. Verify by generating a mixed-type
workbook with both versions and diffing cell by cell; every difference should be
one you can name.

The publish workflow runs from that branch and gives any version below the
current `latest` the `legacy` dist-tag, so a maintenance release cannot become
the default install. Check `npm view xlsx-stream-writer dist-tags` afterwards.

**Dependabot does not watch this branch.** `.github/dependabot.yml` declares no
`target-branch`, so both ecosystems track the default branch only. The coverage
is exactly inverted against the risk: `master` has zero runtime dependencies and
is watched, while `0.2.x` ships `jszip` and `stream-browserify` to the projects
that by definition cannot upgrade away from a problem, and is not. Advisories
against those two will not open a pull request here — check them by hand before
a maintenance release, or add a second pair of `updates` entries with
`target-branch: "0.2.x"` and accept the pull request traffic.


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
  be rolled back after 72 hours, and unpublishing a version burns it forever, so
  never reach for a republish first. Establish whether the release is actually
  sound before doing anything.

  The likeliest cause is timing: npm's attestation endpoint can lag the publish
  by longer than the job is willing to wait, and the message says the
  attestation was unavailable rather than wrong. Fetch it yourself a minute
  later; if it is there, the release is fine and only the check was impatient.
  Confirm the same four things the job does — the provenance subject digest
  matches the tarball npm serves, the repository is this one, the workflow ref
  and path are the ones that ran, and the resolved commit is the release commit.

  A genuine mismatch in any of those is serious and means the artifact on the
  registry is not the one that was built here. A timing failure is not, but it
  leaves `tag-and-release` skipped, so create the tag and the GitHub release by
  hand against the release commit.
- **`tag-and-release` failed.** Cosmetic. Re-run the workflow, or create the tag
  and release by hand; the published artifact is unaffected.
