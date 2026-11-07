using ClosedXML.Excel;

namespace OdcBatchExcel.Tests;

/// <summary>
/// Helpers to build commands and workbooks, run <see cref="OdcBatchExcel.Execute"/> and reopen its output.
/// </summary>
internal static class Batch
{
    public static Command Cmd(string operation, string arguments) => new() { Operation = operation, Arguments = arguments };

    public static ExecuteResult Run(byte[] file, params Command[] commands) => new OdcBatchExcel().Execute(file, [.. commands]);

    public static ExecuteResult Run(params Command[] commands) => Run([], commands);

    // Runs one command on a new workbook and returns the exception it throws.
    public static InvalidOperationException Fail(string operation, string arguments) =>
        Assert.Throws<InvalidOperationException>(() => Run(Cmd(operation, arguments)));

    public static XLWorkbook Open(ExecuteResult result) => new(new MemoryStream(result.File));

    public static byte[] Workbook(Action<XLWorkbook> setup)
    {
        using var workbook = new XLWorkbook();
        setup(workbook);
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    // A workbook whose first sheet, "Data", has A1:B3 filled with a1, b1, a2, ..., and a second sheet, "Other".
    public static byte[] DataWorkbook() => Workbook(workbook =>
    {
        var sheet = workbook.AddWorksheet("Data");
        for (var row = 1; row <= 3; row++)
        {
            sheet.Cell(row, 1).Value = $"a{row}";
            sheet.Cell(row, 2).Value = $"b{row}";
        }

        workbook.AddWorksheet("Other");
    });
}
