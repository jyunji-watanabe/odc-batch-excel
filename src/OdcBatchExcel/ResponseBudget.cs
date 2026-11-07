using System.Globalization;
using System.Text;

namespace OdcBatchExcel;

/// <summary>
/// Keeps one <see cref="IOdcBatchExcel.Execute"/> response under ODC's payload limit, so an
/// oversized batch fails with a message that says what to split instead of a generic platform error.
/// Create one per call; it is not shared.
/// </summary>
/// <remarks>
/// The size is an estimate of the response as ODC sends it: JSON, with binary data as Base64.
/// Read results are counted as their UTF-8 bytes plus the escaping of quotes and backslashes,
/// plus a fixed allowance for each result's field names.
/// </remarks>
internal sealed class ResponseBudget
{
    /// <summary>
    /// ODC's External Logic payload limit (5.5 MB), in bytes.
    /// </summary>
    internal const long DefaultMaxBytes = 5_500_000;

    // {"CommandIndex":2147483647,"Operation":"","Value":""}, plus separators.
    private const int ResultOverheadBytes = 64;

    private readonly long maxBytes;
    private long resultBytes;

    /// <summary>
    /// Initializes a new instance of the <see cref="ResponseBudget"/> class.
    /// </summary>
    /// <param name="maxBytes">The response limit, in bytes.</param>
    public ResponseBudget(long maxBytes)
    {
        this.maxBytes = maxBytes;
    }

    /// <summary>
    /// Counts one read result, and fails as soon as it, or all results so far, exceed the limit.
    /// </summary>
    /// <param name="operation">The result's operation name.</param>
    /// <param name="value">The result's JSON value.</param>
    public void AddResult(string operation, string value)
    {
        var size = EncodedSize(operation) + EncodedSize(value) + ResultOverheadBytes;
        if (size > this.maxBytes)
        {
            throw new InvalidOperationException(
                $"The result is about {Bytes(size)}, more than ODC's response limit of {Bytes(this.maxBytes)}. Read a smaller range.");
        }

        this.resultBytes += size;
        if (this.resultBytes > this.maxBytes)
        {
            throw new InvalidOperationException(
                $"The read results so far are about {Bytes(this.resultBytes)}, more than ODC's response limit of {Bytes(this.maxBytes)}. Split the reads across several Execute calls.");
        }
    }

    /// <summary>
    /// Fails if the returned file plus all read results exceed the limit.
    /// </summary>
    /// <param name="fileLength">The length of the saved workbook, in bytes.</param>
    public void CheckTotal(int fileLength)
    {
        var fileBytes = Base64Size(fileLength);
        var total = fileBytes + this.resultBytes;
        if (total > this.maxBytes)
        {
            throw new InvalidOperationException(
                $"The response would be about {Bytes(total)} (the file as Base64: {Bytes(fileBytes)}; read results: {Bytes(this.resultBytes)}), "
                + $"more than ODC's response limit of {Bytes(this.maxBytes)}. Split the batch across several Execute calls, or make the workbook smaller.");
        }
    }

    /// <summary>
    /// Estimates the bytes a string takes inside a JSON response.
    /// </summary>
    /// <param name="text">The string.</param>
    /// <returns>Its UTF-8 length plus one byte per quote or backslash, which JSON escapes.</returns>
    internal static long EncodedSize(string text)
    {
        return Encoding.UTF8.GetByteCount(text) + text.AsSpan().Count('"') + text.AsSpan().Count('\\');
    }

    private static long Base64Size(long length) => (length + 2) / 3 * 4;

    private static string Bytes(long bytes) => $"{bytes.ToString("N0", CultureInfo.InvariantCulture)} bytes";
}
