using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using ClosedXML.Excel;

namespace OdcBatchExcel.Commands;

/// <summary>
/// Strict parsing of a command's JSON arguments, and validation helpers shared by the operations.
/// Every failure is a <see cref="CommandArgumentException"/> whose message names the argument.
/// </summary>
internal static class CommandArgs
{
    /// <summary>The value types that <c>ValueType</c> and <c>ColumnTypes</c> accept.</summary>
    internal const string ValueTypeNames = "Text, Number, Boolean or DateTime";

    private const string Text = "Text";
    private const string Number = "Number";
    private const string Boolean = "Boolean";
    private const string DateTimeType = "DateTime";

    /// <summary>
    /// Gets the JSON options for arguments and results. They are read-only, so sharing them keeps the library stateless.
    /// </summary>
    internal static JsonSerializerOptions Options { get; } = CreateOptions();

    /// <summary>
    /// Parses <paramref name="arguments"/> into <typeparamref name="T"/>, rejecting unknown, missing,
    /// null and wrongly typed properties with a message that names the property.
    /// </summary>
    /// <typeparam name="T">The operation's argument record.</typeparam>
    /// <param name="arguments">The JSON object from <see cref="Command.Arguments"/>.</param>
    /// <returns>The parsed arguments.</returns>
    internal static T Parse<T>(string? arguments)
        where T : class
    {
        var typeInfo = (JsonTypeInfo<T>)Options.GetTypeInfo(typeof(T));

        if (string.IsNullOrWhiteSpace(arguments))
        {
            throw new CommandArgumentException("Arguments is empty; it must be a JSON object, e.g. {}.");
        }

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(arguments);
        }
        catch (JsonException ex)
        {
            throw new CommandArgumentException($"Arguments is not valid JSON: {ex.Message}", ex);
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                throw new CommandArgumentException($"Arguments must be a JSON object, not {Describe(root.ValueKind)}.");
            }

            CheckProperties(root, typeInfo.Properties);

            try
            {
                return root.Deserialize(typeInfo)
                    ?? throw new CommandArgumentException("Arguments must be a JSON object.");
            }
            catch (JsonException ex)
            {
                throw WrongType(ex, typeInfo.Properties);
            }
        }
    }

    /// <summary>
    /// Serializes a read result to the JSON string stored in <see cref="CommandResult.Value"/>.
    /// </summary>
    /// <typeparam name="T">The result record.</typeparam>
    /// <param name="result">The result.</param>
    /// <returns>The JSON string.</returns>
    internal static string Serialize<T>(T result) => JsonSerializer.Serialize(result, (JsonTypeInfo<T>)Options.GetTypeInfo(typeof(T)));

    /// <summary>
    /// Finds a worksheet by name (case-insensitive). An empty name means the first worksheet.
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="sheet">The worksheet name, or empty.</param>
    /// <returns>The worksheet.</returns>
    internal static IXLWorksheet Sheet(IXLWorkbook workbook, string? sheet)
    {
        if (string.IsNullOrEmpty(sheet))
        {
            return workbook.Worksheets.First();
        }

        if (workbook.TryGetWorksheet(sheet, out var worksheet))
        {
            return worksheet;
        }

        var names = string.Join(", ", workbook.Worksheets.Select(w => $"'{w.Name}'"));
        throw new CommandArgumentException($"Sheet '{sheet}' was not found. Worksheets: {names}.");
    }

    /// <summary>
    /// Checks that <paramref name="address"/> is a single A1-style cell address.
    /// </summary>
    /// <param name="address">The cell address to check.</param>
    /// <param name="property">The argument name, for the error message.</param>
    /// <returns>The same address, once it is known to be valid.</returns>
    internal static string Cell(string address, string property)
    {
        return XLHelper.IsValidA1Address(address)
            ? address
            : throw new CommandArgumentException($"{property} '{address}' is not a valid cell address, e.g. B2.");
    }

    /// <summary>
    /// Checks that <paramref name="address"/> is a cell (<c>B2</c>) or a cell range (<c>A1:D20</c>).
    /// </summary>
    /// <param name="address">The address.</param>
    /// <param name="property">The argument name, for the error message.</param>
    /// <returns>The first and last cell addresses.</returns>
    internal static (string First, string Last) Range(string address, string property)
    {
        var parts = address.Split(':');
        if (parts.Length is 1 or 2 && parts.All(XLHelper.IsValidA1Address))
        {
            return (parts[0], parts[^1]);
        }

        throw new CommandArgumentException($"{property} '{address}' is not a valid range address, e.g. A1:D20.");
    }

    /// <summary>
    /// Converts a column letter (<c>C</c>) to its number, and checks that <paramref name="count"/>
    /// columns starting there fit on the sheet.
    /// </summary>
    /// <param name="column">The column letter.</param>
    /// <param name="count">The number of columns.</param>
    /// <returns>The column number.</returns>
    internal static int Column(string column, int count)
    {
        if (!XLHelper.IsValidColumn(column))
        {
            throw new CommandArgumentException($"Column '{column}' is not a valid column letter, e.g. C.");
        }

        var number = XLHelper.GetColumnNumberFromLetter(column);
        Count(count);
        if (number + count - 1 > XLHelper.MaxColumnNumber)
        {
            throw new CommandArgumentException($"Column {column} + Count {count} goes past the last column, XFD.");
        }

        return number;
    }

    /// <summary>
    /// Parses a column (<c>B</c>) or column span (<c>B:D</c>) into column numbers.
    /// </summary>
    /// <param name="columns">The column or span.</param>
    /// <returns>The first and last column numbers.</returns>
    internal static (int First, int Last) Columns(string columns)
    {
        var parts = columns.Split(':');
        if (parts.Length > 2 || !parts.All(XLHelper.IsValidColumn))
        {
            throw new CommandArgumentException($"Columns '{columns}' is not a valid column or span, e.g. B or B:D.");
        }

        var first = XLHelper.GetColumnNumberFromLetter(parts[0]);
        var last = XLHelper.GetColumnNumberFromLetter(parts[^1]);
        if (first > last)
        {
            throw new CommandArgumentException($"Columns '{columns}' must go from left to right, e.g. B:D.");
        }

        return (first, last);
    }

    /// <summary>
    /// Checks that <paramref name="row"/> is a valid row and that <paramref name="count"/> rows
    /// starting there fit on the sheet.
    /// </summary>
    /// <param name="row">The 1-based row number.</param>
    /// <param name="count">The number of rows.</param>
    internal static void Rows(int row, int count)
    {
        if (row < 1 || row > XLHelper.MaxRowNumber)
        {
            throw new CommandArgumentException($"Row must be between 1 and {XLHelper.MaxRowNumber}, but it is {row}.");
        }

        Count(count);
        if (row + count - 1 > XLHelper.MaxRowNumber)
        {
            throw new CommandArgumentException($"Row {row} + Count {count} goes past the last row, {XLHelper.MaxRowNumber}.");
        }
    }

    /// <summary>
    /// Normalizes a value type name. Empty means Text.
    /// </summary>
    /// <param name="valueType">The value type name (case-insensitive).</param>
    /// <param name="property">The argument name, for the error message.</param>
    /// <returns>The canonical value type name.</returns>
    internal static string ValueType(string? valueType, string property)
    {
        if (string.IsNullOrEmpty(valueType))
        {
            return Text;
        }

        return Array.Find([Text, Number, Boolean, DateTimeType], name => string.Equals(valueType, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new CommandArgumentException($"{property} '{valueType}' is not valid. Use {ValueTypeNames}.");
    }

    /// <summary>
    /// Converts a string to a cell value of the given type. An empty string is a blank cell.
    /// </summary>
    /// <param name="value">The value as a string.</param>
    /// <param name="valueType">The canonical value type from <see cref="ValueType"/>.</param>
    /// <param name="property">The argument name, for the error message.</param>
    /// <returns>The cell value.</returns>
    internal static XLCellValue Value(string value, string valueType, string property)
    {
        if (value.Length == 0)
        {
            return Blank.Value;
        }

        switch (valueType)
        {
            case Number:
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
                {
                    return number;
                }

                throw new CommandArgumentException($"{property} '{value}' is not a valid Number. Use the invariant format, e.g. 1234.5.");

            case Boolean:
                if (bool.TryParse(value, out var boolean))
                {
                    return boolean;
                }

                throw new CommandArgumentException($"{property} '{value}' is not a valid Boolean. Use true or false.");

            case DateTimeType:
                if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var dateTime))
                {
                    return dateTime;
                }

                throw new CommandArgumentException($"{property} '{value}' is not a valid DateTime. Use ISO 8601, e.g. 2026-11-07T09:30:00.");

            default:
                return value;
        }
    }

    private static JsonSerializerOptions CreateOptions()
    {
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
            RespectNullableAnnotations = true,
            RespectRequiredConstructorParameters = true,
            TypeInfoResolver = new DefaultJsonTypeInfoResolver(),
        };
        options.MakeReadOnly();
        return options;
    }

    private static void CheckProperties(JsonElement root, IList<JsonPropertyInfo> properties)
    {
        var present = new HashSet<string>(StringComparer.Ordinal);
        foreach (var json in root.EnumerateObject())
        {
            var property = properties.FirstOrDefault(p => string.Equals(p.Name, json.Name, StringComparison.OrdinalIgnoreCase))
                ?? throw new CommandArgumentException($"Unknown property '{json.Name}'. {Allowed(properties)}");

            if (json.Value.ValueKind == JsonValueKind.Null && !property.IsSetNullable)
            {
                throw new CommandArgumentException($"Property '{property.Name}' must not be null.");
            }

            present.Add(property.Name);
        }

        var missing = properties.Where(p => p.IsRequired && !present.Contains(p.Name)).Select(p => $"'{p.Name}'").ToList();
        if (missing.Count == 1)
        {
            throw new CommandArgumentException($"Missing required property {missing[0]}.");
        }

        if (missing.Count > 1)
        {
            throw new CommandArgumentException($"Missing required properties {string.Join(", ", missing)}.");
        }
    }

    private static CommandArgumentException WrongType(JsonException ex, IList<JsonPropertyInfo> properties)
    {
        // The path looks like "$.Count" or "$.Rows[0][1]"; its first segment is the property.
        var path = ex.Path ?? string.Empty;
        var name = path.StartsWith("$.", StringComparison.Ordinal) ? path[2..] : string.Empty;
        var end = name.IndexOfAny(['[', '.']);
        if (end >= 0)
        {
            name = name[..end];
        }

        var property = properties.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));

        return property is null
            ? new CommandArgumentException($"Arguments could not be read: {ex.Message}", ex)
            : new CommandArgumentException($"Property '{property.Name}' must be {Describe(property.PropertyType)}.", ex);
    }

    private static string Allowed(IList<JsonPropertyInfo> properties)
    {
        return properties.Count == 0
            ? "This operation takes no arguments; use {}."
            : $"Allowed properties: {string.Join(", ", properties.Select(p => p.Name))}.";
    }

    private static string Describe(Type type)
    {
        if (type == typeof(string))
        {
            return "a string";
        }

        if (type == typeof(int))
        {
            return "a whole number";
        }

        if (type == typeof(double))
        {
            return "a number";
        }

        if (type == typeof(IReadOnlyList<string>))
        {
            return "an array of strings";
        }

        return type == typeof(IReadOnlyList<IReadOnlyList<string>>)
            ? "an array of arrays of strings"
            : "a valid value";
    }

    private static string Describe(JsonValueKind kind) => kind switch
    {
        JsonValueKind.Array => "an array",
        JsonValueKind.String => "a string",
        JsonValueKind.Number => "a number",
        JsonValueKind.True or JsonValueKind.False => "a boolean",
        _ => "null",
    };

    private static void Count(int count)
    {
        if (count < 1)
        {
            throw new CommandArgumentException($"Count must be 1 or more, but it is {count}.");
        }
    }
}
