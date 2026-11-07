using OutSystems.ExternalLibraries.SDK;

namespace OdcBatchExcel;

/// <summary>
/// One operation to run on the workbook, such as setting a cell value or reading a range.
/// </summary>
[OSStructure(Description = "One operation to run on the workbook, such as setting a cell value or reading a range.")]
public struct Command
{
    /// <summary>
    /// Gets or sets the operation name, for example <c>SetCellValue</c>. Case-insensitive.
    /// </summary>
    [OSStructureField(Description = "The operation name, for example SetCellValue. Case-insensitive.", IsMandatory = true)]
    public string Operation { get; set; }

    /// <summary>
    /// Gets or sets the operation's arguments as a JSON object, for example <c>{"Cell":"B2","Value":"42"}</c>.
    /// </summary>
    [OSStructureField(Description = "The operation's arguments as a JSON object. Its shape depends on Operation.", IsMandatory = true)]
    public string Arguments { get; set; }
}
