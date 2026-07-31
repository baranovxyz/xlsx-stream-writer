# Security policy

## Supported versions

This package has a single maintainer. Two lines are published.

| Version | Supported |
| ------- | --------- |
| 1.x     | yes — active, on the `latest` tag |
| 0.2.7   | frozen — on the `legacy` tag; a patch is published if one is warranted |
| < 0.2.7 | no — deprecated on npm; silently writes wrong cell values |

Frozen means no proactive work: nothing watches that branch between releases,
and a report against it will take longer to answer. It does not mean abandoned.
A fix for input that produces a wrong or unopenable file is backported and
published as a `0.2.z`; a fix that would change output which already worked is
not, because callers on that line pinned it for stability.

Everything below 0.2.7 is deprecated, but **not for a security reason**, and the
distinction is worth keeping straight. Those versions look up shared strings in
a plain object, so a cell holding `constructor` or `__proto__` finds an
inherited member instead of missing and writes a reference that is not an index.
The cell then reads back as a different string, or the workbook fails to open.

That is a data-correctness bug. It pollutes nothing — the assignment is single
level, and `__proto__ = <number>` is ignored — injects nothing, and crosses no
trust boundary. Nor did installing those versions expose a caller to a
vulnerable dependency: the `jszip` range floats, so a fresh install resolves the
fixed version regardless of what the shipped lockfile pinned.

## Reporting a vulnerability

Report privately through
[GitHub Security Advisories](https://github.com/baranovxyz/xlsx-stream-writer/security/advisories/new).
Please do not open a public issue for an unpatched vulnerability.

Include a reproduction if you can — for this package that usually means the
input rows, options, and the resulting `.xlsx` or XML.

Expect an acknowledgement within a week.

## Threat model

This package turns caller-supplied values into XML inside a ZIP container. The
security-relevant surface is small but real:

- **XML injection.** Cell values, shared strings and style attributes are
  escaped before they reach the document. A value that escapes its element or
  attribute would let a caller inject arbitrary markup into a workbook that
  another party opens.
- **Malformed output.** Characters that are illegal in XML 1.0 produce a file
  Excel refuses to open, or silently repairs by discarding content.
- **Resource use.** Row and column counts, and shared-string cardinality, are
  driven by caller input. Callers streaming untrusted data should bound it.

This package only *writes* spreadsheets. It never parses untrusted `.xlsx`
input, so zip-parsing classes of bug — path traversal on extract, zip bombs —
do not apply.

## Supply chain

- Published through npm OIDC trusted publishing from a GitHub Actions workflow,
  with SLSA provenance attached. No maintainer holds a long-lived npm token.
- The published tarball is packed and tested before the credential-holding job
  ever runs, and is pinned by SHA-512 integrity between the two.
- Dependencies are audited at `--audit-level=low` on every push.

Verify a release yourself with:

```sh
npm view xlsx-stream-writer dist.attestations
npm audit signatures
```
