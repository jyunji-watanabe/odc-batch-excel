namespace OdcBatchExcel;

/// <summary>
/// Implementation of <see cref="IOdcBatchExcel"/>. ODC creates a new instance per call,
/// so do not keep state between calls.
/// </summary>
public sealed class OdcBatchExcel : IOdcBatchExcel
{
    /// <inheritdoc/>
    public string Greet(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        return $"Hello, {name.Trim()}!";
    }
}
