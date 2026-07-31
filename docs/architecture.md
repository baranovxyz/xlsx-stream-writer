# Architecture

Design decisions that the code does not explain on its own. Read this before
changing the ZIP writer, the stream plumbing, or the order parts are written in.

## Async iterables are the only stream abstraction

Node streams, web streams, generators and plain arrays disagree on nearly
everything, but they can all be reduced to an async iterable — and that
reduction needs no polyfill in either environment. Every input is normalised to
one on the way in, and everything internal is an async generator.

This is what let the package drop its stream compatibility dependency. It also
means backpressure is inherent: nothing is produced until a consumer pulls.

Entry point: `src/streams.ts`.

## Part order is a contract, not a convention

The shared-string table is populated *as the worksheet is walked*. It is only
complete once the last row has been converted, so the worksheet part must be
written before it.

Previously this held by accident — the zip library happened to consume entries
in insertion order. It is now guaranteed by the writer, which consumes each
entry's source only when that entry's turn comes.

**The trap:** a web `ReadableStream` starts pulling the moment it is
constructed. Handing one to the archive writer emits the shared-string header —
including its counts — before the sheet has been walked, producing a table
declaring zero strings that nevertheless lists them. Entry sources are therefore
async generators, which do nothing until iterated. Do not "simplify" this by
reusing the public stream properties.

Entry point: the part-assembly method in `src/xlsx-stream-writer.ts`.

## The writer walks its rows exactly once

Rows arrive as a stream, and a stream can only be consumed once. Both the
public XML stream properties and the workbook builders consume the same rows, so
using one closes off the other. Attempting both raises rather than producing a
workbook that is silently empty.

## Each archive entry is buffered, then decided

Two shapes exist for a ZIP entry, and the choice cannot be revisited once the
entry's header is written:

- **Exact sizes in the header.** The most widely readable form, but it requires
  knowing the compressed size and checksum up front.
- **Sizes deferred to a trailing descriptor.** Necessary when the content is
  still streaming, and the only form that stays correct past the format's 4 GiB
  ceiling, which requires the ZIP64 records.

The writer compresses into memory until the entry either finishes — in which
case it takes the first shape — or crosses a size threshold, at which point it
commits to the second and streams the rest. Small workbooks, which is nearly all
of them, get maximally compatible output; large ones stay correct.

The threshold is on *uncompressed* bytes, deliberately. ZIP64 is required when
the uncompressed size overflows, and spreadsheet XML compresses roughly tenfold,
so a compressed-size threshold would miss the case it exists for.

Entry point: `src/zip/writer.ts`.

## Compression is swapped at bundle time, not runtime

Node and browsers both compress, through entirely different APIs. Rather than
branch at runtime, the package ships two adapters with the same shape and lets
the bundler substitute one for the other through the `browser` field in
`package.json`.

The consequence worth knowing: a stray Node builtin import anywhere else in the
sources would break every browser build, because only the one adapter is
redirected. A test enforces that no other built module reaches for one.

The Node adapter is preferred where both would work, because it exposes the
compression level and the browser API does not.

## Archives are reproducible

Entry timestamps are fixed rather than taken from the clock, so the same rows
always produce the same bytes. This makes byte-exact golden tests possible and
makes releases verifiable. Do not "fix" the timestamps to be current.

## Excel's grid is validated up front

A workbook that exceeds Excel's row or column limits opens as a repair prompt
that discards content. The writer raises at the offending row instead, naming
it. The same reasoning applies to values with no representation in the format:
non-finite numbers and characters XML cannot encode become blank or are removed
rather than written through.
