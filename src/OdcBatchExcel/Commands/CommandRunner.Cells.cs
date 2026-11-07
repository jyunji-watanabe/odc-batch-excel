using ClosedXML.Excel;

namespace OdcBatchExcel.Commands;

/// <content>
/// Write operations for cell values and formulas.
/// </content>
internal static partial class CommandRunner
{
    private static string? SetCellValue(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<SetCellValueArgs>(arguments);
        var cell = CommandArgs.Cell(args.Cell, "Cell");
        var valueType = CommandArgs.ValueType(args.ValueType, "ValueType");
        var value = CommandArgs.Value(args.Value, valueType, "Value");

        CommandArgs.Sheet(workbook, args.Sheet).Cell(cell).Value = value;
        return null;
    }

    private static string? SetFormula(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<SetFormulaArgs>(arguments);
        var cell = CommandArgs.Cell(args.Cell, "Cell");
        var formula = args.Formula.StartsWith('=') ? args.Formula[1..] : args.Formula;
        if (string.IsNullOrWhiteSpace(formula))
        {
            throw new CommandArgumentException("Formula is empty.");
        }

        CommandArgs.Sheet(workbook, args.Sheet).Cell(cell).FormulaA1 = formula;
        return null;
    }

    private static string? SetRangeValues(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<SetRangeValuesArgs>(arguments);
        var startCell = CommandArgs.Cell(args.StartCell, "StartCell");
        if (args.Rows.Count == 0)
        {
            throw new CommandArgumentException("Rows is empty; it needs at least one row.");
        }

        var columnTypes = ColumnTypes(args);
        var sheet = CommandArgs.Sheet(workbook, args.Sheet);
        var start = sheet.Cell(startCell);
        var lastRow = start.Address.RowNumber + args.Rows.Count - 1;
        var lastColumn = start.Address.ColumnNumber + args.Rows.Max(r => r?.Count ?? 0) - 1;
        if (lastRow > XLHelper.MaxRowNumber || lastColumn > XLHelper.MaxColumnNumber)
        {
            throw new CommandArgumentException($"Rows starting at {startCell} go past the edge of the sheet.");
        }

        // Convert everything first, so a bad value leaves the sheet unchanged.
        var values = new List<XLCellValue[]>(args.Rows.Count);
        for (var r = 0; r < args.Rows.Count; r++)
        {
            var row = args.Rows[r] ?? throw new CommandArgumentException($"Rows[{r}] must not be null.");
            var converted = new XLCellValue[row.Count];
            for (var c = 0; c < row.Count; c++)
            {
                var property = $"Rows[{r}][{c}]";
                var value = row[c] ?? throw new CommandArgumentException($"{property} must not be null.");
                converted[c] = CommandArgs.Value(value, columnTypes?[c] ?? "Text", property);
            }

            values.Add(converted);
        }

        for (var r = 0; r < values.Count; r++)
        {
            for (var c = 0; c < values[r].Length; c++)
            {
                start.CellBelow(r).CellRight(c).Value = values[r][c];
            }
        }

        return null;
    }

    private static string[]? ColumnTypes(SetRangeValuesArgs args)
    {
        if (args.ColumnTypes is null)
        {
            return null;
        }

        var types = args.ColumnTypes
            .Select((t, i) => CommandArgs.ValueType(t, $"ColumnTypes[{i}]"))
            .ToArray();

        for (var r = 0; r < args.Rows.Count; r++)
        {
            var count = args.Rows[r]?.Count ?? 0;
            if (count != types.Length)
            {
                throw new CommandArgumentException(
                    $"Rows[{r}] has {count} values, but ColumnTypes has {types.Length}. Every row must have one value per column type.");
            }
        }

        return types;
    }
}
