# Adversarial review

A recurring pass that assumes the suite is green and proves less than it appears
to. Two have run over this package, and both shipped a release: 1.2.0 and 1.3.0.
This is how the next one starts from what the last two learned instead of from
nothing.

## Why it keeps paying

Not one finding was a crash. Every one produced a workbook that looked fine:

- A cell holding the string `"constructor"` corrupted the entire workbook. The
  shared-string table was a plain object, so the value found an inherited member
  of `Object.prototype` where an index belonged. The sheet then referenced
  `<v>function Object() { [native code] }</v>` and the string table came out
  empty. Present in **every version published before 1.2.0**, `0.2.6` included.
- The same collision in style `format` and `fill` values.
- Reading `sheetXmlStream` and then calling `getFile()` produced a silently empty
  workbook: rows can only be walked once, and nothing said so.
- Reading `sharedStringsXmlStream` before the worksheet produced a part
  declaring a count of zero and listing nothing — valid-looking and quietly
  wrong.
- `styleIdFunc` return values reached the `s` attribute unvalidated, so a string
  could close the attribute and open another.

A test suite that inspects what the writer generates agrees with the writer.
None of the above moves a single existing assertion, which is the point: an
adversarial pass is looking for what the suite cannot see, not for what it
already checks.

## The shapes that keep recurring

Start here. Each of these has produced a real bug in this package, and the
conventions in [AGENTS.md](../AGENTS.md) are the fossilised remains of them.

1. **A caller-supplied string used as a lookup key.** Anything reaching a plain
   object as a key. Two unrelated tables have been hit.
2. **State that can only be consumed once.** Generators, streams, and rows.
   Ask what happens if a caller reads it twice, or reads part B before part A.
3. **A value from a caller's callback that reaches XML.** It is caller-supplied
   even though it arrived by return rather than by argument.
4. **Module-level mutable state aliased into an instance.** The second workbook
   built in a process inherits the first one's state.
5. **A part serialised before the data that fills it.** Order of emission is
   load-bearing, and getting it wrong yields a well-formed, empty part.

New findings belong on this list.

## Running one

Read the code trying to answer one question: *what input produces a workbook
that is wrong, and that the suite calls correct?* Prefer output corruption over
crashes — a crash is already loud, and this package's failure mode is silence.

Then:

- **Every finding gets a regression test**, in `tests/robustness/` unless
  something else fits better, and a changelog entry. A fix without a test is an
  invitation to a third pass finding the same thing.
- **Decide which lines it lands on.** Output that is wrong or unopenable usually
  belongs on `0.2.x` as well as `master`; anything that changes output which
  already worked belongs on `master` alone. See [AGENTS.md](../AGENTS.md).
- **Add the shape above** if it is one this list would not have caught.

Worth running before a release that changes behaviour, and after any phase of
work large enough to have its own changelog section.

## What it does not cover

An adversarial pass reasons about the code. It cannot tell you a real reader
accepts the output — for that see
[docs/verifying-output.md](verifying-output.md), which records what has actually
been opened and what has not.
