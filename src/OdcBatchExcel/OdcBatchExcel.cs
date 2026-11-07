using ClosedXML.Excel;
using OdcBatchExcel.Commands;

namespace OdcBatchExcel;

/// <summary>
/// Implementation of <see cref="IOdcBatchExcel"/>. ODC creates a new instance per call,
/// so do not keep state between calls.
/// </summary>
public sealed class OdcBatchExcel : IOdcBatchExcel
{
    /// <summary>
    /// The sheet name a new, empty workbook starts with.
    /// </summary>
    internal const string DefaultSheetName = "Sheet1";

    private readonly long maxResponseBytes;

    /// <summary>
    /// Initializes a new instance of the <see cref="OdcBatchExcel"/> class. ODC uses this constructor.
    /// </summary>
    public OdcBatchExcel()
        : this(ResponseBudget.DefaultMaxBytes)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="OdcBatchExcel"/> class with a custom response limit, for tests.
    /// </summary>
    /// <param name="maxResponseBytes">The response limit, in bytes.</param>
    internal OdcBatchExcel(long maxResponseBytes)
    {
        this.maxResponseBytes = maxResponseBytes;
    }

    /// <inheritdoc/>
    public ExecuteResult Execute(byte[] file, List<Command> commands)
    {
        ArgumentNullException.ThrowIfNull(commands);

        using var input = new MemoryStream(file ?? [], writable: false);
        using var workbook = input.Length == 0 ? CreateWorkbook() : OpenWorkbook(input);

        var budget = new ResponseBudget(this.maxResponseBytes);
        var results = new List<CommandResult>();
        for (var i = 0; i < commands.Count; i++)
        {
            var index = i + 1;
            var (operation, value) = RunCommand(workbook, commands[i], index, budget);
            if (value is not null)
            {
                results.Add(new CommandResult { CommandIndex = index, Operation = operation, Value = value });
            }
        }

        // ClosedXML saves formulas without their calculated values, so ask Excel (and other
        // spreadsheet apps) to recalculate everything when the file is opened.
        workbook.FullCalculationOnLoad = true;

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        var bytes = output.ToArray();
        budget.CheckTotal(bytes.Length);
        return new ExecuteResult { File = bytes, Results = results };
    }

    private static (string Operation, string? Value) RunCommand(XLWorkbook workbook, Command command, int index, ResponseBudget budget)
    {
        try
        {
            var (operation, value) = CommandRunner.Run(workbook, command.Operation, command.Arguments);
            if (value is not null)
            {
                budget.AddResult(operation, value);
            }

            return (operation, value);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            throw new InvalidOperationException($"Command #{index} ({command.Operation}): {ex.Message}", ex);
        }
    }

    private static XLWorkbook CreateWorkbook()
    {
        var workbook = new XLWorkbook();
        workbook.AddWorksheet(DefaultSheetName);
        return workbook;
    }

    private static XLWorkbook OpenWorkbook(MemoryStream input)
    {
        try
        {
            return new XLWorkbook(input);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            throw new ArgumentException($"File is not a valid .xlsx workbook: {ex.Message}", nameof(input), ex);
        }
    }
}
