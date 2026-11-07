namespace OdcBatchExcel.Commands;

// Argument records, one per operation (or per group of operations with the same arguments).
// System.Text.Json creates them; `required` marks required properties, and nullable
// properties are optional. Property names are the JSON names (case-insensitive).

/// <summary>Arguments of <c>SetCellValue</c>.</summary>
internal sealed record SetCellValueArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the cell address, e.g. <c>B2</c>.</summary>
    public required string Cell { get; init; }

    /// <summary>Gets the value as a string; empty clears the cell.</summary>
    public required string Value { get; init; }

    /// <summary>Gets how to interpret <see cref="Value"/>: Text (default), Number, Boolean or DateTime.</summary>
    public string? ValueType { get; init; }
}

/// <summary>Arguments of <c>SetFormula</c>.</summary>
internal sealed record SetFormulaArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the cell address, e.g. <c>B2</c>.</summary>
    public required string Cell { get; init; }

    /// <summary>Gets the A1-style formula; the leading <c>=</c> is optional.</summary>
    public required string Formula { get; init; }
}

/// <summary>Arguments of <c>SetRangeValues</c>.</summary>
internal sealed record SetRangeValuesArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the top-left cell of the block, e.g. <c>A2</c>.</summary>
    public required string StartCell { get; init; }

    /// <summary>Gets the rows of values to write.</summary>
    public required IReadOnlyList<IReadOnlyList<string>> Rows { get; init; }

    /// <summary>Gets the value type of each column; omitted means every column is Text.</summary>
    public IReadOnlyList<string>? ColumnTypes { get; init; }
}

/// <summary>Arguments of <c>InsertRows</c> and <c>DeleteRows</c>.</summary>
internal sealed record RowsArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the 1-based row number.</summary>
    public required int Row { get; init; }

    /// <summary>Gets the number of rows.</summary>
    public int Count { get; init; } = 1;
}

/// <summary>Arguments of <c>InsertColumns</c> and <c>DeleteColumns</c>.</summary>
internal sealed record ColumnsArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the column letter, e.g. <c>C</c>.</summary>
    public required string Column { get; init; }

    /// <summary>Gets the number of columns.</summary>
    public int Count { get; init; } = 1;
}

/// <summary>Arguments of <c>SetColumnWidth</c>.</summary>
internal sealed record SetColumnWidthArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the column or column span, e.g. <c>B</c> or <c>B:D</c>.</summary>
    public required string Columns { get; init; }

    /// <summary>Gets the width in Excel character units.</summary>
    public required double Width { get; init; }
}

/// <summary>Arguments of <c>AutoFitColumns</c>.</summary>
internal sealed record AutoFitColumnsArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the column or column span; empty means every used column.</summary>
    public string? Columns { get; init; }
}

/// <summary>Arguments of <c>GetCellValue</c>.</summary>
internal sealed record CellArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the cell address, e.g. <c>B2</c>.</summary>
    public required string Cell { get; init; }
}

/// <summary>Arguments of <c>GetRangeValues</c>.</summary>
internal sealed record RangeArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }

    /// <summary>Gets the range address, e.g. <c>A1:D20</c>.</summary>
    public required string Range { get; init; }
}

/// <summary>Arguments of <c>GetUsedRange</c>.</summary>
internal sealed record SheetArgs
{
    /// <summary>Gets the worksheet name; empty means the first worksheet.</summary>
    public string? Sheet { get; init; }
}

/// <summary>Arguments of <c>GetSheetNames</c>: none, so only <c>{}</c> is accepted.</summary>
internal sealed record NoArgs;

/// <summary>Result of <c>GetCellValue</c>.</summary>
/// <param name="Value">The value as a string.</param>
/// <param name="Type">Blank, Text, Number, Boolean, DateTime, TimeSpan or Error.</param>
internal sealed record CellValueResult(string Value, string Type);

/// <summary>Result of <c>GetRangeValues</c>.</summary>
/// <param name="Rows">The values, row by row.</param>
internal sealed record RangeValuesResult(IReadOnlyList<IReadOnlyList<string>> Rows);

/// <summary>Result of <c>GetUsedRange</c>.</summary>
/// <param name="Address">The used range, e.g. <c>A1:D20</c>; empty for an empty sheet.</param>
/// <param name="FirstRow">The first used row, or 0.</param>
/// <param name="FirstColumn">The first used column number, or 0.</param>
/// <param name="RowCount">The number of rows, or 0.</param>
/// <param name="ColumnCount">The number of columns, or 0.</param>
internal sealed record UsedRangeResult(string Address, int FirstRow, int FirstColumn, int RowCount, int ColumnCount);

/// <summary>Result of <c>GetSheetNames</c>.</summary>
/// <param name="Names">The worksheet names, in workbook order.</param>
internal sealed record SheetNamesResult(IReadOnlyList<string> Names);
