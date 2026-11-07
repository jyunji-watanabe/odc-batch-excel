using ClosedXML.Excel;

namespace OdcBatchExcel.Commands;

/// <content>
/// Write operations for rows and columns.
/// </content>
internal static partial class CommandRunner
{
    // Excel's maximum column width, in character units.
    private const double MaxColumnWidth = 255;

    private static string? InsertRows(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<RowsArgs>(arguments);
        CommandArgs.Rows(args.Row, args.Count);

        CommandArgs.Sheet(workbook, args.Sheet).Row(args.Row).InsertRowsAbove(args.Count);
        return null;
    }

    private static string? DeleteRows(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<RowsArgs>(arguments);
        CommandArgs.Rows(args.Row, args.Count);

        CommandArgs.Sheet(workbook, args.Sheet).Rows(args.Row, args.Row + args.Count - 1).Delete();
        return null;
    }

    private static string? InsertColumns(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<ColumnsArgs>(arguments);
        var column = CommandArgs.Column(args.Column, args.Count);

        CommandArgs.Sheet(workbook, args.Sheet).Column(column).InsertColumnsBefore(args.Count);
        return null;
    }

    private static string? DeleteColumns(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<ColumnsArgs>(arguments);
        var column = CommandArgs.Column(args.Column, args.Count);

        CommandArgs.Sheet(workbook, args.Sheet).Columns(column, column + args.Count - 1).Delete();
        return null;
    }

    private static string? SetColumnWidth(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<SetColumnWidthArgs>(arguments);
        var (first, last) = CommandArgs.Columns(args.Columns);
        if (args.Width is <= 0 or > MaxColumnWidth)
        {
            throw new CommandArgumentException($"Width must be more than 0 and at most {MaxColumnWidth}, but it is {args.Width}.");
        }

        CommandArgs.Sheet(workbook, args.Sheet).Columns(first, last).Width = args.Width;
        return null;
    }

    private static string? AutoFitColumns(IXLWorkbook workbook, string? arguments)
    {
        var args = CommandArgs.Parse<AutoFitColumnsArgs>(arguments);
        var sheet = CommandArgs.Sheet(workbook, args.Sheet);

        if (string.IsNullOrEmpty(args.Columns))
        {
            sheet.ColumnsUsed().AdjustToContents();
        }
        else
        {
            var (first, last) = CommandArgs.Columns(args.Columns);
            sheet.Columns(first, last).AdjustToContents();
        }

        return null;
    }
}
