using System.Globalization;
using ClosedXML.Excel;

namespace OdcBatchExcel.Commands;

/// <content>
/// Read operations. Each returns its result as JSON; formula cells return their calculated value.
/// </content>
internal static partial class CommandRunner
{
    // Keeps a single read well inside ODC's 5.5 MB payload limit and the 1 GB memory limit.
    private const int MaxRangeCells = 1_000_000;

    private static string? GetCellValue(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<CellArgs>(arguments);
        var cell = CommandArgs.Cell(args.Cell, "Cell");

        var (value, type) = Read(CommandArgs.Sheet(workbook, args.Sheet).Cell(cell));
        return CommandArgs.Serialize(new CellValueResult(value, type));
    }

    private static string? GetRangeValues(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<RangeArgs>(arguments);
        var (first, last) = CommandArgs.Range(args.Range, "Range");
        var range = CommandArgs.Sheet(workbook, args.Sheet).Range(first, last);

        var rowCount = range.RowCount();
        var columnCount = range.ColumnCount();
        if ((long)rowCount * columnCount > MaxRangeCells)
        {
            throw new CommandArgumentException($"Range '{args.Range}' has more than {MaxRangeCells:N0} cells; read it in smaller parts.");
        }

        var rows = new List<IReadOnlyList<string>>(rowCount);
        for (var r = 1; r <= rowCount; r++)
        {
            var row = new string[columnCount];
            for (var c = 1; c <= columnCount; c++)
            {
                row[c - 1] = Read(range.Cell(r, c)).Value;
            }

            rows.Add(row);
        }

        return CommandArgs.Serialize(new RangeValuesResult(rows));
    }

    private static string? GetUsedRange(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<SheetArgs>(arguments);
        var used = CommandArgs.Sheet(workbook, args.Sheet).RangeUsed(XLCellsUsedOptions.Contents);

        if (used is null)
        {
            return CommandArgs.Serialize(new UsedRangeResult(string.Empty, 0, 0, 0, 0));
        }

        var first = used.RangeAddress.FirstAddress;
        var last = used.RangeAddress.LastAddress;
        return CommandArgs.Serialize(new UsedRangeResult(
            $"{first.ColumnLetter}{first.RowNumber}:{last.ColumnLetter}{last.RowNumber}",
            first.RowNumber,
            first.ColumnNumber,
            used.RowCount(),
            used.ColumnCount()));
    }

    private static string? GetSheetNames(IXLWorkbook workbook, string? arguments)
    {
        CommandArgs.Parse<NoArgs>(arguments);
        return CommandArgs.Serialize(new SheetNamesResult([.. workbook.Worksheets.Select(w => w.Name)]));
    }

    private static (string Value, string Type) Read(IXLCell cell)
    {
        XLCellValue value;
        try
        {
            value = cell.Value;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            throw new InvalidOperationException($"Could not calculate the formula in {cell.Address}: {ex.Message}", ex);
        }

        var invariant = CultureInfo.InvariantCulture;
        return value.Type switch
        {
            XLDataType.Text => (value.GetText(), "Text"),
            XLDataType.Number => (value.GetNumber().ToString(invariant), "Number"),
            XLDataType.Boolean => (value.GetBoolean() ? "true" : "false", "Boolean"),
            XLDataType.DateTime => (value.GetDateTime().ToString("yyyy-MM-ddTHH:mm:ss.FFFFFFF", invariant), "DateTime"),
            XLDataType.TimeSpan => (value.GetTimeSpan().ToString("c", invariant), "TimeSpan"),
            XLDataType.Error => (ErrorText(value.GetError()), "Error"),
            _ => (string.Empty, "Blank"),
        };
    }

    private static string ErrorText(XLError error) => error switch
    {
        XLError.NullValue => "#NULL!",
        XLError.DivisionByZero => "#DIV/0!",
        XLError.IncompatibleValue => "#VALUE!",
        XLError.CellReference => "#REF!",
        XLError.NameNotRecognized => "#NAME?",
        XLError.NumberInvalid => "#NUM!",
        XLError.NoValueAvailable => "#N/A",
        _ => error.ToString(),
    };
}
