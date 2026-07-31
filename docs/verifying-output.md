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

## What a spreadsheet application has actually read

Everything above establishes that the archive is well-formed and that the XML
parses. None of it proves a spreadsheet application is happy, so that is checked
separately, by `tests/archive/spreadsheet.test.js`.

It builds four workbooks, has LibreOffice convert them headless to CSV, and
compares every cell against what was written:

- **Both string paths** — the shared table and inline strings — must round-trip
  exactly, Cyrillic included.
- **A sheet in the streamed entry shape**: 120 000 rows, past the writer's
  buffering limit, so the local header sizes are zero and the real sizes trail
  the data. The test asserts that shape before reading, so the archive form most
  likely to be rejected is provably the one being read.
- **A fixture of values chosen to break readers**: XML metacharacters, a string
  that tries to close its attribute and open an element, `__proto__` /
  `constructor` / `toString` / `valueOf` / `hasOwnProperty` as cell values,
  formula- and DDE-shaped strings, leading and trailing whitespace, emoji, RTL,
  combining marks, CJK, a 32 767-character string, and numeric and boolean edge
  cases. Formula-shaped strings must come back as text rather than evaluated,
  and the `Object.prototype` names as themselves.

First run, on 2026-07-31 against 1.3.0 with LibreOffice Calc 24.2.7.2: every
cell round-tripped, and separately a 500 000-row workbook converted with all
500 000 rows intact. No file triggered a repair prompt.

These checks **skip silently when LibreOffice is absent**, so the suite stays
portable — which also means they pass while asserting nothing. CI installs
`libreoffice-calc` and sets `XLSX_REQUIRE_SPREADSHEET=1`, which turns a missing
LibreOffice into a failure, so a runner that loses it cannot go quietly green.
Set it locally to prove the checks are really running.

Note that a LibreOffice installation without its spreadsheet component will fail
to load *any* spreadsheet — including a plain CSV — with a generic "source file
could not be loaded". That failure looks like a broken file and is not; install
`libreoffice-calc` specifically, not `libreoffice`, before concluding anything
about the output.

## What is still unverified

**Real Excel.** LibreOffice is a proxy, not the thing itself; it is more
tolerant than Excel in places. Before relying on a release, open a generated
workbook in real Excel once.

How tolerant, measured rather than assumed: given a shared-string reference that
is not a number — the shape 0.2.6 wrote for a cell holding `constructor` —
LibreOffice 24.2.7.2 does not refuse the file and does not prompt for repair. It
coerces the reference to index 0 and prints the first string in the table.
Surrounding cells keep their correct indices, so nothing looks disturbed.

That is the failure mode this proxy is worst at catching: the CSV comes back
with a *plausible wrong value* rather than an error, and a cell-by-cell
comparison only catches it because the expected value is known. A corruption
that happened to substitute the value a test also expected would pass. Green
here means "LibreOffice read it and agreed with us", never "the file is
well-formed" — the readers above are what establish that.

Nothing here covers styling beyond what the examples exercise, and the CSV proxy
flattens formatting by design — a value that reads back correctly says nothing
about how it is displayed.
