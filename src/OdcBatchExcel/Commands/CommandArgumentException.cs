namespace OdcBatchExcel.Commands;

/// <summary>
/// A command's arguments are missing, malformed or out of range.
/// </summary>
/// <remarks>
/// <see cref="IOdcBatchExcel.Execute"/> wraps it in an <see cref="InvalidOperationException"/> that names the command.
/// </remarks>
public sealed class CommandArgumentException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="CommandArgumentException"/> class.
    /// </summary>
    public CommandArgumentException()
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="CommandArgumentException"/> class.
    /// </summary>
    /// <param name="message">A message that names the offending argument.</param>
    public CommandArgumentException(string message)
        : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="CommandArgumentException"/> class.
    /// </summary>
    /// <param name="message">A message that names the offending argument.</param>
    /// <param name="innerException">The underlying error.</param>
    public CommandArgumentException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
