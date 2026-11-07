# OdcBatchExcel design

## Purpose

ODC apps call External Logic actions remotely, one call at a time. Editing an Excel file
operation by operation means one round trip, and one upload and download of the file, per
operation. OdcBatchExcel takes a workbook and a **list of commands**, runs them all in one call,
and returns the updated workbook along with anything the commands read.

## Architecture

| # | Part | Where | Status |
|---|---|---|---|
| 1 | External Logic library (C#, .NET 10, ClosedXML) | `src/OdcBatchExcel/` | Built |
| 2 | Unit tests (xUnit v3) | `tests/OdcBatchExcel.Tests/` | Built |
| 3 | ODC Library: one Action per operation that builds a `Command`, plus `Execute` | ODC tenant | Not built yet; see [ODC Library](#part-3-odc-library-not-built-yet) |
| 4 | ODC test app using BDDFramework | ODC tenant | Not built yet; see [BDD test app](#part-4-bdd-test-app-not-built-yet) |

The Excel library is [ClosedXML](https://github.com/ClosedXML/ClosedXML) (MIT license). It reads
and writes .xlsx, calculates formulas, and is built on Microsoft's OpenXML SDK. The packaged zip is
about 3 MB.

## Contract

The library exposes a single action. It must stay stable: renaming an action or changing a
signature breaks the apps that use it.

```
Execute(File: Binary Data, Commands: List of Command) -> Result: ExecuteResult

Command        { Operation: Text, Arguments: Text }                  // Arguments is a JSON object
ExecuteResult  { File: Binary Data, Results: List of CommandResult }
CommandResult  { CommandIndex: Integer, Operation: Text, Value: Text } // Value is a JSON object
```

- **Input file:** if `File` is empty, Execute starts from a new workbook with one sheet,
  `Sheet1`. Otherwise `File` must be an .xlsx workbook.
- **Order:** commands run in list order on one in-memory workbook. Each command sees the
  changes made by the commands before it, so a batch can write values and formulas and then
  read the calculated results.
- **Operation names** are case-insensitive. Results always use the canonical name.
- **Reads:** each read command adds one `CommandResult`, in command order. `CommandIndex` is
  the command's 1-based position. Write commands add no results.
- **All or nothing:** if any command fails, the whole action fails and no file is returned.
  The error message starts with the command's position and the operation as it was given, e.g.
  `Command #3 (SetCellValue): Value '1,5' is not a valid Number. ...`.
- **Recalculation:** the returned file asks Excel to recalculate fully when it opens
  (`fullCalcOnLoad`). See [Known limitations](#known-limitations).

The ODC contract is generic on purpose. A new operation only needs a new name and its own
JSON arguments, so the External Logic signature never changes.

## Operations

Every operation except `GetSheetNames` takes an optional `Sheet`, the worksheet name
(case-insensitive). If it's empty or left out, the operation uses the first worksheet.

### Writes

| Operation | Arguments (**bold** = required) | Notes |
|---|---|---|
| `SetCellValue` | **`Cell`**, **`Value`**, `ValueType` | `Cell` is A1-style (`B2`). `ValueType` is `Text` (default), `Number`, `Boolean` or `DateTime`. |
| `SetFormula` | **`Cell`**, **`Formula`** | A1-style formula. The leading `=` is optional. |
| `SetRangeValues` | **`StartCell`**, **`Rows`**, `ColumnTypes` | `Rows` is an array of arrays of strings. Without `ColumnTypes`, all values are text and rows may have different lengths. With it, every row must have exactly one value per column type. |
| `InsertRows` | **`Row`**, `Count` (1) | Inserts above `Row`. Data and formula references below move down. |
| `DeleteRows` | **`Row`**, `Count` (1) | Data below moves up. |
| `InsertColumns` | **`Column`**, `Count` (1) | `Column` is a letter (`C`). Inserts to the left of it. |
| `DeleteColumns` | **`Column`**, `Count` (1) | |
| `SetColumnWidth` | **`Columns`**, **`Width`** | `Columns` is `B` or `B:D`. `Width` is in Excel character units, more than 0 and at most 255. |
| `AutoFitColumns` | `Columns` | Left empty, it fits every used column. |

**Value conversion** (used by `SetCellValue` and `SetRangeValues`):
- **Empty string:** always clears the cell.
- **Number:** uses the invariant format (`1234.5`, `-1e3`). `NaN` and infinities are rejected.
- **Boolean:** `true` or `false`, case-insensitive.
- **DateTime:** ISO 8601 (`2026-11-07` or `2026-11-07T09:30:00`). A value with an offset or `Z`
  is converted to UTC. One without is stored as given. Excel has no time zones.

A header row and typed data rows need two `SetRangeValues` commands, one for the header (text)
and one for the data with `ColumnTypes`, because `ColumnTypes` applies to every row.

### Reads

All values in results are strings. Numbers use the invariant culture, dates are
`yyyy-MM-ddTHH:mm:ss` (fractional seconds only when present) and booleans are `true`/`false`.
A formula cell returns its calculated value.

| Operation | Arguments | Result `Value` |
|---|---|---|
| `GetCellValue` | **`Cell`** | `{"Value":"42","Type":"Number"}`. `Type` is one of `Blank`, `Text`, `Number`, `Boolean`, `DateTime`, `TimeSpan` or `Error`. For errors, `Value` is the Excel error code, e.g. `#DIV/0!`. |
| `GetRangeValues` | **`Range`** | `{"Rows":[["a","1"],["b","2"]]}`. `Range` is `A1:D20` or a single cell. Whole-column ranges like `A:A` aren't accepted. At most 1,000,000 cells. |
| `GetUsedRange` | — | `{"Address":"A1:D20","FirstRow":1,"FirstColumn":1,"RowCount":20,"ColumnCount":4}`. On an empty sheet, `Address` is `""` and the numbers are `0`. Only cells with content count, not formatting alone. |
| `GetSheetNames` | none: `{}` | `{"Names":["Sheet1","Data"]}`, in workbook order. |

## Argument rules

Arguments are parsed strictly, so a mistake in the calling app fails loudly instead of being
silently ignored:

- `Arguments` must be a JSON object. Even an operation without arguments needs `{}`.
- Property names are case-insensitive.
- These are all errors, each with a message that names the property:
  - an unknown property, including one that belongs to a different operation
  - a missing required property
  - `null` in a property
  - a value of the wrong JSON type (e.g. `"Count":"2"`)
- After parsing, values are checked: addresses must be valid, `Row`/`Count` ≥ 1, everything
  must fit on the sheet (1,048,576 rows, columns up to `XFD`), and `ValueType` must be a known
  type.
- `SetRangeValues` converts every value before it writes anything, so a bad value leaves the
  sheet unchanged.

## Limits

- **ODC External Logic:** payloads up to 5.5 MB in each direction (the file is sent both ways),
  95 seconds per call and 1 GB of memory. Split very large workbooks or reads across calls.
- **Reads:** `GetRangeValues` refuses ranges of more than 1,000,000 cells.

## Known limitations

These come from ClosedXML. The first three are pinned by tests, so a ClosedXML upgrade that
changes them is noticed.

- **Deleted references:** when a row or column that a formula references is deleted, Excel
  shows `#REF!`, but ClosedXML points the reference at the neighboring row/column instead
  (`DeleteRows_ReferenceToADeletedRow_MovesToTheRowAbove_UnlikeExcel`).
- **Unsupported formulas:** ClosedXML can't calculate every Excel feature, e.g. the range
  intersection operator (`SUM(A1:A2 B1:B2)`). Writing such a formula works, but reading its value
  fails with `Could not calculate the formula in A1: ...`. Excel calculates it when the file is
  opened.
- **No cached formula values:** saved formulas don't include their calculated values. Excel
  recalculates on open (`fullCalcOnLoad`). Tools that read cached values without calculating,
  such as file previews or pandas/openpyxl in values-only mode, see those cells as empty.
- **AutoFit:** `AutoFitColumns` measures text with fonts available to ClosedXML. Widths in the
  ODC Linux container may differ slightly from local results. Check it when trying the library in
  an app.
- **File format:** only .xlsx. Legacy .xls files are not supported.

## Code layout

```
src/OdcBatchExcel/
  IOdcBatchExcel.cs              [OSInterface] contract: the Execute action
  OdcBatchExcel.cs               Execute: open/create the workbook, run each command, save
  Command.cs, ExecuteResult.cs,
  CommandResult.cs               [OSStructure] types
  Commands/
    CommandRunner.cs             Operation catalog (name, argument record, handler) and dispatch
    CommandRunner.Cells.cs       SetCellValue, SetFormula, SetRangeValues
    CommandRunner.RowsColumns.cs InsertRows, DeleteRows, InsertColumns, DeleteColumns,
                                 SetColumnWidth, AutoFitColumns
    CommandRunner.Reads.cs       GetCellValue, GetRangeValues, GetUsedRange, GetSheetNames
    Arguments.cs                 One record per operation's arguments, plus the read result records
    CommandArgs.cs               Strict JSON parsing and shared validation (sheets, addresses, values)
    CommandArgumentException.cs  Thrown for bad arguments; Execute wraps it with the command index
tests/OdcBatchExcel.Tests/
  ContractTests.cs               The ODC-facing contract (names, signature, structure fields)
  ExecuteTests.cs                Behavior of every operation, batching and general errors
  ArgumentValidationTests.cs     Valid/invalid argument tables for every operation, plus completeness checks
  Batch.cs                       Test helpers
```

Each operation is a private static method that parses its arguments with
`CommandArgs.Parse<TArgs>`, validates them, and changes or reads the workbook. It returns `null`
for a write or the result JSON for a read. The catalog and the JSON options are immutable, which
keeps the library stateless as ODC requires.

### Adding an operation

1. Add an argument record in `Commands/Arguments.cs`. Mark required properties `required` and
   make optional ones nullable or give them a default. For a read, also add a result record.
2. Write the private method in the right `CommandRunner.*.cs` file, and add it to the catalog in
   `CommandRunner.cs` with its argument record type.
3. Add rows to `ArgumentValidationTests`: minimal and full valid arguments, a missing-property
   case for each required property, and an unknown-property case. The completeness tests fail
   until these exist. Add behavior tests to `ExecuteTests`.
4. Update the operation tables in this document. In ODC, add the matching command Action to the
   ODC Library (part 3).

The `Execute` signature does not change, so apps that already use the library keep working.

## Part 3: ODC Library (not built yet)

The ODC Library hides the JSON from app developers:

- **One Action per operation**, e.g. `Cmd_SetCellValue(Sheet, Cell, Value, ValueType)`. Each takes
  typed inputs, fills a structure with the operation's argument names, serializes it with
  `JSONSerialize`, and returns a `Command` with the operation name. Leave optional inputs out of
  the JSON when they are empty, or send `""` for `Sheet`, which means the first sheet.
- **An `Execute` Action** that wraps the External Logic `Execute`.
- **One typed helper per read**, e.g. `Result_GetCellValue(CommandResult) -> {Value, Type}`.
  It uses `JSONDeserialize` into a structure with the result shape above.

An app builds a list of commands with these Actions, calls `Execute` once, saves the returned
file, and reads the typed results.

## Part 4: BDD test app (not built yet)

An ODC app that uses BDDFramework to test the ODC Library end to end, with at least one scenario
per operation. A scenario builds commands with the library Actions, runs `Execute` on a small
workbook, and checks the results through the read helpers. Scenarios should also cover one
failing batch, to confirm that the error message (`Command #n (Operation): ...`) reaches the app,
and the AutoFit widths in the ODC container.
