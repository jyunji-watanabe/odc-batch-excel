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
    /// Returns a greeting for the given name.
    /// </summary>
    /// <param name="name">The name to greet.</param>
    /// <returns>A greeting message.</returns>
    [OSAction(Description = "Returns a greeting for the given name.", ReturnName = "Greeting")]
    string Greet([OSParameter(Description = "The name to greet.")] string name);
}
