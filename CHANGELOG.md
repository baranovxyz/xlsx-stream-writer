# Changelog

All notable changes to this package are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [semantic versioning](https://semver.org/spec/v2.0.0.html).

## 0.2.7 — 2026-07-31

Security and tooling only. No API or output changes: the generated workbook is
byte-for-byte identical to 0.2.6.

### Security

- Cleared all 59 known advisories (11 critical, 19 high) reported by
  `npm audit` against the 0.2.6 dependency tree.
- `jszip` moved from `^3.1.5` to `^3.10.1`, picking up fixes for the
  prototype-pollution and path-traversal advisories affecting the pinned 3.1.5.
- `stream-browserify` moved from `^2.0.2` to `^3.0.0`.
- Removed `jest@25` and the unused `crc` devDependency, taking the development
  dependency tree from 576 packages to zero. Tests now run on the built-in
  `node:test` runner.
- Regenerated `package-lock.json` at lockfile v3; the previous v1 lockfile was
  why the resolved tree stayed four years stale.

### Added

- Continuous integration across Node 20.19, 22 and 24, with `npm audit` at
  `--audit-level=low` and a full-history secret scan.
- Release pipeline publishing through npm OIDC trusted publishing with SLSA
  provenance, including registry integrity and provenance verification after
  publish. See `docs/RELEASING.md`.
- Test suite now inspects the generated `.xlsx` archive itself — entry order,
  per-part bytes, XML well-formedness, CRCs and streaming-mode flags — rather
  than only in-memory streams.
- `LICENSE` file. The package has always declared MIT but shipped no license
  text.

### Changed

- `npm test` runs once and exits, so it works in CI. Use `npm run test:watch`
  for the previous watching behaviour.
- The published tarball now contains only `index.js`, `src/`, `README.md`,
  `CHANGELOG.md` and `LICENSE`. Previous releases also shipped `tests/`,
  `examples/` and the lockfile.
- Declared `engines.node` as `>=20.19.0`.

### Removed

- `examples/crc_speed_test.js`, a scratch benchmark that depended on the
  now-removed `crc` package and had a latent bug of its own.

## 0.2.6 and earlier

See the git history.
