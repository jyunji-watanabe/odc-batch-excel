using OutSystems.ExternalLibraries.SDK;

namespace OdcBatchExcel;

/// <summary>
/// The output of one read command.
/// </summary>
[OSStructure(Description = "The output of one read command.")]
public struct CommandResult
{
    /// <summary>
    /// Gets or sets the 1-based index of the command in the command list.
    /// </summary>
    [OSStructureField(Description = "The 1-based index of the command in the command list.")]
    public int CommandIndex { get; set; }

    /// <summary>
    /// Gets or sets the operation name, for example <c>GetCellValue</c>.
    /// </summary>
    [OSStructureField(Description = "The operation name, for example GetCellValue.")]
    public string Operation { get; set; }

    /// <summary>
    /// Gets or sets the result as a JSON object. Its shape depends on the operation.
    /// </summary>
    [OSStructureField(Description = "The result as a JSON object. Its shape depends on the operation.")]
    public string Value { get; set; }
}
