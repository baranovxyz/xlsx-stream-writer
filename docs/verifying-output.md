# Verifying output

How this package establishes that the workbooks it writes are actually valid,
and what that stops short of. Read this before treating a green suite as proof
that a file opens.

## The problem

Excel does not report *why* a workbook is malformed. It reports "unreadable
content" and offers to repair the file, which discards data. A test suite that
only inspects the strings the package generates will happily pass while
producing files nothing can open — that is exactly the state this package was in
before it was hardened.

So correctness is established by reading the produced bytes back, with readers
that do not share the writer's assumptions.

## Three independent readers

Every archive assertion goes through at least one reader that could disagree
with the writer:

1. **The suite's own ZIP parser** (`tests/support/unzip.js`). Reads the central
   directory rather than walking local headers, so it handles archives written
   in streaming mode — the shape where local headers carry zeroes and the real
   sizes trail the data. It verifies every entry's checksum and length while
   parsing, so a successful parse is itself an assertion.

2. **A third-party ZIP library**, kept as a development dependency for exactly
   this purpose. It shares no lineage with the writer, so a blind spot in one is
   unlikely to be a blind spot in both. This is the *only* reason that
   dependency still exists; it is not a runtime dependency and must not become
   one.

3. **The system `unzip`**, when present. Skipped silently when it is not, so the
   suite stays portable.

Anything the suite's own parser would reject is something a real reader may
reject too.

## Byte-exact goldens

The expected content of every part is pinned byte for byte, captured from the
implementation before the hardening work began. This is what makes it possible
to refactor aggressively — replacing the ZIP layer, then converting to
TypeScript — while proving the output never moved.

A diff against a golden means the behaviour changed. That belongs in the
changelog, not in an updated fixture.

## Branches no fixture can reach

Some code only runs on inputs too large to build in a test: the archive records
that appear past 4 GiB, and past roughly 65 thousand entries. Rather than leave
those untested, the record builders are exported under an internal-only name and
driven directly with the sizes that trigger them. The package's export map does
not expose that module, so consumers cannot reach it.

The threshold that switches an entry to its streaming shape is also injectable,
so that path is exercised with small fixtures instead of multi-megabyte ones —
and separately, once, with a workbook genuinely large enough to cross the real
threshold.

## What is still unverified

**No workbook produced by this package has been opened in Excel or LibreOffice
during its hardening.** Everything above establishes that the archive is
well-formed and that the XML parses; none of it proves Excel is happy.

LibreOffice is the practical proxy, via a headless conversion. Note that a
LibreOffice installation without its spreadsheet component will fail to load
*any* spreadsheet — including a plain CSV — with a generic "source file could
not be loaded". That failure looks like a broken file and is not; check that the
spreadsheet component is actually installed before concluding anything about the
output.

Before relying on a release, open a generated workbook in real Excel once.
