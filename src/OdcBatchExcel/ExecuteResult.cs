using OutSystems.ExternalLibraries.SDK;

namespace OdcBatchExcel;

/// <summary>
/// The output of <see cref="IOdcBatchExcel.Execute"/>.
/// </summary>
[OSStructure(Description = "The workbook after all commands, plus one result per read command.")]
public struct ExecuteResult
{
    /// <summary>
    /// Gets or sets the workbook after all commands, as .xlsx bytes.
    /// </summary>
    [OSStructureField(Description = "The workbook after all commands, as .xlsx.")]
    public byte[] File { get; set; }

    /// <summary>
    /// Gets or sets one result per read command, in command order. Write commands add none.
    /// </summary>
    [OSStructureField(Description = "One result per read command, in command order. Write commands add none.")]
    public List<CommandResult> Results { get; set; }
}
