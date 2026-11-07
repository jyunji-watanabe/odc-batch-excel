using System.Collections.Frozen;
using System.Collections.Immutable;
using ClosedXML.Excel;

namespace OdcBatchExcel.Commands;

/// <summary>
/// Maps operation names to the private methods that run them. Each operation lives in one
/// partial file by area: cell values, rows and columns, and reads.
/// </summary>
internal static partial class CommandRunner
{
    // To add an operation, write its private method and argument record, then list it here.
    // Both fields are immutable, so the library stays stateless. Catalog must come first:
    // static fields initialize in textual order.
    private static readonly ImmutableArray<Operation> Catalog =
    [
        new("SetCellValue", typeof(SetCellValueArgs), SetCellValue),
        new("SetFormula", typeof(SetFormulaArgs), SetFormula),
        new("SetRangeValues", typeof(SetRangeValuesArgs), SetRangeValues),
        new("InsertRows", typeof(RowsArgs), InsertRows),
        new("DeleteRows", typeof(RowsArgs), DeleteRows),
        new("InsertColumns", typeof(ColumnsArgs), InsertColumns),
        new("DeleteColumns", typeof(ColumnsArgs), DeleteColumns),
        new("SetColumnWidth", typeof(SetColumnWidthArgs), SetColumnWidth),
        new("AutoFitColumns", typeof(AutoFitColumnsArgs), AutoFitColumns),
        new("GetCellValue", typeof(CellArgs), GetCellValue),
        new("GetRangeValues", typeof(RangeArgs), GetRangeValues),
        new("GetUsedRange", typeof(SheetArgs), GetUsedRange),
        new("GetSheetNames", typeof(NoArgs), GetSheetNames),
    ];

    private static readonly FrozenDictionary<string, Operation> ByName =
        Catalog.ToFrozenDictionary(o => o.Name, StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Gets every supported operation, in catalog order.
    /// </summary>
    internal static ImmutableArray<Operation> Operations => Catalog;

    /// <summary>
    /// Runs one command on <paramref name="workbook"/>.
    /// </summary>
    /// <param name="workbook">The workbook to change or read.</param>
    /// <param name="operation">The operation name (case-insensitive).</param>
    /// <param name="arguments">The operation's JSON arguments.</param>
    /// <returns>The canonical operation name, and the result JSON for a read or <see langword="null"/> for a write.</returns>
    internal static (string Operation, string? Value) Run(IXLWorkbook workbook, string? operation, string? arguments)
    {
        if (string.IsNullOrWhiteSpace(operation))
        {
            throw new CommandArgumentException("Operation is empty.");
        }

        if (!ByName.TryGetValue(operation, out var found))
        {
            var supported = string.Join(", ", Operations.Select(o => o.Name));
            throw new CommandArgumentException($"Unknown operation '{operation}'. Supported: {supported}.");
        }

        return (found.Name, found.Handler(workbook, arguments));
    }

    /// <summary>
    /// One operation: its canonical name, its argument record and the method that runs it.
    /// </summary>
    /// <param name="Name">The canonical operation name.</param>
    /// <param name="ArgumentsType">The record its JSON arguments are parsed into.</param>
    /// <param name="Handler">Runs the operation; returns result JSON for a read, or <see langword="null"/> for a write.</param>
    internal sealed record Operation(string Name, Type ArgumentsType, Func<IXLWorkbook, string?, string?> Handler);
}
