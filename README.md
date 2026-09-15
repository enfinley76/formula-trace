# formula-trace

You inherit a spreadsheet. Someone left three years ago. Cell `C1` has a
number in it and you need to know: if I change `A1`, does that number move?
And if I want to understand why `C1` is what it is, which cells do I need to
read first?

Excel's own "trace precedents" arrows are fine for one hop but fall apart
once formulas are nested a few levels deep, and they only work inside the
app itself. `formula-trace` does the same thing from the command line, on a
plain CSV export, and walks the whole chain instead of one hop.

It answers exactly one question: given a cell, what is the full set of cells
feeding into it (or depending on it), in the order they'd need to be
computed.

## Input format

A CSV with a header row of `cell,formula`. A formula cell starts with `=`;
anything else is treated as a literal value.

```csv
cell,formula
A1,10
A2,20
A3,5
B1,=A1+A2
B2,=B1*2
B3,=B2-A3
C1,=B2+A1
C2,=C1/A3
```

This is what you get from "Save As CSV" after temporarily switching a sheet
to show formulas instead of values, or from a quick export script against
the `.xlsx` XML.

## Usage

What does `C2` depend on, in calculation order?

```
$ formula-trace testdata/example.csv C2
A1	(literal)
A2	(literal)
B1	=A1+A2
B2	=B1*2
A3	(literal)
C1	=B2+A1
```

That's the order you'd need to fill in values to hand-check `C2` by hand, or
the order a recalculation engine would need to visit them in.

What breaks if `A1` changes?

```
$ formula-trace -dependents testdata/example.csv A1
B1	=A1+A2
B2	=B1*2
B3	=B2-A3
C1	=B2+A1
C2	=C1/A3
```

If the sheet has a circular reference, the tool reports the cycle instead of
hanging:

```
$ formula-trace testdata/circular.csv A1
error: circular reference: A1 -> B1 -> A1
```

## Multiple sheets

There's no separate "sheet" column. If your export covers more than one
sheet, prefix the cell in the `cell` column with the sheet name, the same
way a formula would reference it: `Sheet2!A1`. Formulas can reference other
sheets the same way, including ranges (`SUM(Sheet1!A1:A2)`); the sheet
qualifier only needs to appear once, before the first cell of a range.

```csv
cell,formula
Sheet1!A1,10
Sheet1!A2,20
Sheet1!B1,=Sheet1!A1+Sheet1!A2
Sheet2!A1,=Sheet1!B1*2
```

If a formula references a cell that isn't in the CSV at all - commonly
because the sheet it lives on wasn't included in the export - it still
shows up in the trace, labeled `(not in sheet)` rather than `(literal)`, so
you know the chain is incomplete rather than assuming it ends there.

## Building

Standard library only, no dependencies.

```
go build -o formula-trace .
```

## Known limitations

- No named ranges, no `INDIRECT`, no functions are evaluated — this traces
  references, not values.
- Columns beyond `ZZZ` haven't been tested.
- Sheet names with an apostrophe in them can't be quoted correctly, since
  the parser looks for a plain `'...'` wrapper.
