using OutSystems.ExternalLibraries.SDK;

namespace OdcBatchExcel;

/// <summary>
/// The ODC External Logic contract. Every method here becomes a server action in ODC.
/// </summary>
/// <remarks>
/// <see cref="OSInterfaceAttribute.Name"/> is the library's identity in the tenant:
/// keep it stable, or ODC will treat the upload as a brand-new library.
/// </remarks>
[OSInterface(Name = "OdcBatchExcel", Description = "Batch processing of Excel files for ODC apps.")]
public interface IOdcBatchExcel
{
    /// <summary>
    /// Runs a batch of commands on an Excel workbook in one call, in list order.
    /// </summary>
    /// <param name="file">The .xlsx workbook to process. Empty starts a new workbook with one sheet, "Sheet1".</param>
    /// <param name="commands">The commands to run, in order. Each one has an operation name and its JSON arguments.</param>
    /// <returns>The workbook after all commands, plus one result per read command.</returns>
    /// <exception cref="InvalidOperationException">A command failed. The message names its 1-based index and operation; no file is returned.</exception>
    [OSAction(
        Description = "Runs a batch of commands on an Excel workbook in one call, in list order. If any command fails, the action fails and no file is returned.",
        ReturnName = "Result",
        ReturnDescription = "The workbook after all commands, plus one result per read command.")]
    ExecuteResult Execute(
        [OSParameter(Description = "The .xlsx workbook to process. Empty starts a new workbook with one sheet, \"Sheet1\".")] byte[] file,
        [OSParameter(Description = "The commands to run, in order. Each one has an operation name and its JSON arguments.")] List<Command> commands);
}
