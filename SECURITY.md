# Security policy

## Supported versions

Only the latest published version receives fixes. This package has a single
maintainer; there are no backport branches.

| Version | Supported |
| ------- | --------- |
| 0.2.7+  | yes       |
| < 0.2.7 | no — depends on `jszip@3.1.5`, which carries known advisories |

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
