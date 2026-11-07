using OdcBatchExcel.Commands;
using static OdcBatchExcel.Tests.Batch;

namespace OdcBatchExcel.Tests;

/// <summary>
/// Every operation takes its own JSON arguments. These tables check, per operation, that missing,
/// unmatched (unknown, wrongly typed, null) and out-of-range arguments are rejected with a message
/// that names the argument, and that minimal and full valid arguments are accepted.
/// </summary>
public class ArgumentValidationTests
{
    /// <summary>
    /// Gets (operation, arguments) pairs that must succeed on a new, empty workbook: the minimal and
    /// the full set of arguments for each operation.
    /// </summary>
    public static TheoryData<string, string> ValidCases { get; } = new()
    {
        { "SetCellValue", """{"Cell":"B2","Value":"x"}""" },
        { "SetCellValue", """{"Sheet":"Sheet1","Cell":"B2","Value":"42","ValueType":"Number"}""" },
        { "SetCellValue", """{"sheet":"sheet1","cell":"b2","value":"x","valuetype":"text"}""" },
        { "SetCellValue", """{"Sheet":null,"Cell":"B2","Value":"x","ValueType":null}""" },
        { "SetFormula", """{"Cell":"A1","Formula":"1+1"}""" },
        { "SetFormula", """{"Sheet":"Sheet1","Cell":"A1","Formula":"=1+1"}""" },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["a"]]}""" },
        { "SetRangeValues", """{"Sheet":"Sheet1","StartCell":"A1","Rows":[["a","1"],["b","2"]],"ColumnTypes":["Text","Number"]}""" },
        { "InsertRows", """{"Row":1}""" },
        { "InsertRows", """{"Sheet":"Sheet1","Row":1048574,"Count":3}""" },
        { "DeleteRows", """{"Row":1}""" },
        { "DeleteRows", """{"Sheet":"Sheet1","Row":2,"Count":3}""" },
        { "InsertColumns", """{"Column":"A"}""" },
        { "InsertColumns", """{"Sheet":"Sheet1","Column":"XFB","Count":3}""" },
        { "DeleteColumns", """{"Column":"A"}""" },
        { "DeleteColumns", """{"Sheet":"Sheet1","Column":"b","Count":2}""" },
        { "SetColumnWidth", """{"Columns":"A","Width":20}""" },
        { "SetColumnWidth", """{"Sheet":"Sheet1","Columns":"A:C","Width":255}""" },
        { "AutoFitColumns", "{}" },
        { "AutoFitColumns", """{"Sheet":"Sheet1","Columns":"A:B"}""" },
        { "GetCellValue", """{"Cell":"A1"}""" },
        { "GetCellValue", """{"Sheet":"Sheet1","Cell":"$A$1"}""" },
        { "GetRangeValues", """{"Range":"A1"}""" },
        { "GetRangeValues", """{"Sheet":"Sheet1","Range":"A1:C3"}""" },
        { "GetUsedRange", "{}" },
        { "GetUsedRange", """{"Sheet":"Sheet1"}""" },
        { "GetSheetNames", "{}" },
    };

    /// <summary>
    /// Gets (operation, arguments, expected message) cases that must fail. The message is what
    /// follows "Command #1 (Operation): " and may be a prefix of the full message.
    /// </summary>
    public static TheoryData<string, string, string> InvalidCases { get; } = new()
    {
        // SetCellValue
        { "SetCellValue", """{"Value":"1"}""", "Missing required property 'Cell'." },
        { "SetCellValue", """{"Cell":"A1"}""", "Missing required property 'Value'." },
        { "SetCellValue", """{"Sheet":"Sheet1"}""", "Missing required properties 'Cell', 'Value'." },
        { "SetCellValue", """{"Cel":"A1","Value":"1"}""", "Unknown property 'Cel'. Allowed properties: Sheet, Cell, Value, ValueType." },
        { "SetCellValue", """{"Cell":"A1","Value":"1","Formula":"1+1"}""", "Unknown property 'Formula'." },
        { "SetCellValue", """{"Cell":"A1","Value":1}""", "Property 'Value' must be a string." },
        { "SetCellValue", """{"Cell":2,"Value":"1"}""", "Property 'Cell' must be a string." },
        { "SetCellValue", """{"Cell":"A1","Value":"1","ValueType":1}""", "Property 'ValueType' must be a string." },
        { "SetCellValue", """{"Cell":"A1","Value":"1","Sheet":["Sheet1"]}""", "Property 'Sheet' must be a string." },
        { "SetCellValue", """{"Cell":null,"Value":"1"}""", "Property 'Cell' must not be null." },
        { "SetCellValue", """{"Cell":"A1","Value":null}""", "Property 'Value' must not be null." },
        { "SetCellValue", """{"Cell":"1A","Value":"1"}""", "Cell '1A' is not a valid cell address" },
        { "SetCellValue", """{"Cell":"A1:B2","Value":"1"}""", "Cell 'A1:B2' is not a valid cell address" },
        { "SetCellValue", """{"Cell":"","Value":"1"}""", "Cell '' is not a valid cell address" },
        { "SetCellValue", """{"Cell":"A1","Value":"1","ValueType":"Integer"}""", "ValueType 'Integer' is not valid. Use Text, Number, Boolean or DateTime." },
        { "SetCellValue", """{"Cell":"A1","Value":"1,5","ValueType":"Number"}""", "Value '1,5' is not a valid Number." },
        { "SetCellValue", """{"Cell":"A1","Value":"NaN","ValueType":"Number"}""", "Value 'NaN' is not a valid Number." },
        { "SetCellValue", """{"Cell":"A1","Value":"yes","ValueType":"Boolean"}""", "Value 'yes' is not a valid Boolean." },
        { "SetCellValue", """{"Cell":"A1","Value":"tomorrow","ValueType":"DateTime"}""", "Value 'tomorrow' is not a valid DateTime." },
        { "SetCellValue", """{"Sheet":"Nope","Cell":"A1","Value":"1"}""", "Sheet 'Nope' was not found." },

        // SetFormula
        { "SetFormula", """{"Formula":"1+1"}""", "Missing required property 'Cell'." },
        { "SetFormula", """{"Cell":"A1"}""", "Missing required property 'Formula'." },
        { "SetFormula", """{"Cell":"A1","Value":"1+1"}""", "Unknown property 'Value'." },
        { "SetFormula", """{"Cell":"A1","Formula":5}""", "Property 'Formula' must be a string." },
        { "SetFormula", """{"Cell":"A1","Formula":null}""", "Property 'Formula' must not be null." },
        { "SetFormula", """{"Cell":"A1","Formula":"="}""", "Formula is empty." },
        { "SetFormula", """{"Cell":"A1","Formula":"  "}""", "Formula is empty." },
        { "SetFormula", """{"Cell":"ZZZZ1","Formula":"1"}""", "Cell 'ZZZZ1' is not a valid cell address" },

        // SetRangeValues
        { "SetRangeValues", """{"Rows":[["a"]]}""", "Missing required property 'StartCell'." },
        { "SetRangeValues", """{"StartCell":"A1"}""", "Missing required property 'Rows'." },
        { "SetRangeValues", """{"StartCell":"A1","Values":[["a"]]}""", "Unknown property 'Values'." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["a"]],"ValueType":"Text"}""", "Unknown property 'ValueType'." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":"a,b"}""", "Property 'Rows' must be an array of arrays of strings." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":["a","b"]}""", "Property 'Rows' must be an array of arrays of strings." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["a",1]]}""", "Property 'Rows' must be an array of arrays of strings." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":null}""", "Property 'Rows' must not be null." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[]}""", "Rows is empty; it needs at least one row." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[null]}""", "Rows[0] must not be null." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["a",null]]}""", "Rows[0][1] must not be null." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["a"]],"ColumnTypes":"Text"}""", "Property 'ColumnTypes' must be an array of strings." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["a"]],"ColumnTypes":["Txt"]}""", "ColumnTypes[0] 'Txt' is not valid." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["a","b"],["c"]],"ColumnTypes":["Text","Text"]}""", "Rows[1] has 1 values, but ColumnTypes has 2." },
        { "SetRangeValues", """{"StartCell":"A1","Rows":[["1,5"]],"ColumnTypes":["Number"]}""", "Rows[0][0] '1,5' is not a valid Number." },
        { "SetRangeValues", """{"StartCell":"XFD1","Rows":[["a","b"]]}""", "Rows starting at XFD1 go past the edge of the sheet." },
        { "SetRangeValues", """{"StartCell":"A1048576","Rows":[["a"],["b"]]}""", "Rows starting at A1048576 go past the edge of the sheet." },

        // InsertRows
        { "InsertRows", """{"Count":1}""", "Missing required property 'Row'." },
        { "InsertRows", """{"Row":1,"Column":"A"}""", "Unknown property 'Column'." },
        { "InsertRows", """{"Row":"2"}""", "Property 'Row' must be a whole number." },
        { "InsertRows", """{"Row":1.5}""", "Property 'Row' must be a whole number." },
        { "InsertRows", """{"Row":1,"Count":"two"}""", "Property 'Count' must be a whole number." },
        { "InsertRows", """{"Row":null}""", "Property 'Row' must not be null." },
        { "InsertRows", """{"Row":1,"Count":null}""", "Property 'Count' must not be null." },
        { "InsertRows", """{"Row":0}""", "Row must be between 1 and 1048576, but it is 0." },
        { "InsertRows", """{"Row":1048577}""", "Row must be between 1 and 1048576, but it is 1048577." },
        { "InsertRows", """{"Row":1,"Count":0}""", "Count must be 1 or more, but it is 0." },
        { "InsertRows", """{"Row":1,"Count":-1}""", "Count must be 1 or more, but it is -1." },
        { "InsertRows", """{"Row":1048576,"Count":2}""", "Row 1048576 + Count 2 goes past the last row, 1048576." },

        // DeleteRows
        { "DeleteRows", "{}", "Missing required property 'Row'." },
        { "DeleteRows", """{"Rows":1}""", "Unknown property 'Rows'." },
        { "DeleteRows", """{"Row":true}""", "Property 'Row' must be a whole number." },
        { "DeleteRows", """{"Row":-3}""", "Row must be between 1 and 1048576, but it is -3." },
        { "DeleteRows", """{"Row":2,"Count":0}""", "Count must be 1 or more, but it is 0." },

        // InsertColumns
        { "InsertColumns", """{"Count":1}""", "Missing required property 'Column'." },
        { "InsertColumns", """{"Column":"A","Row":1}""", "Unknown property 'Row'." },
        { "InsertColumns", """{"Column":3}""", "Property 'Column' must be a string." },
        { "InsertColumns", """{"Column":"3"}""", "Column '3' is not a valid column letter, e.g. C." },
        { "InsertColumns", """{"Column":"XFE"}""", "Column 'XFE' is not a valid column letter, e.g. C." },
        { "InsertColumns", """{"Column":"A1"}""", "Column 'A1' is not a valid column letter, e.g. C." },
        { "InsertColumns", """{"Column":"XFD","Count":2}""", "Column XFD + Count 2 goes past the last column, XFD." },
        { "InsertColumns", """{"Column":"A","Count":0}""", "Count must be 1 or more, but it is 0." },

        // DeleteColumns
        { "DeleteColumns", "{}", "Missing required property 'Column'." },
        { "DeleteColumns", """{"Columns":"A"}""", "Unknown property 'Columns'." },
        { "DeleteColumns", """{"Column":""}""", "Column '' is not a valid column letter, e.g. C." },
        { "DeleteColumns", """{"Column":"A","Count":1.0e1}""", "Property 'Count' must be a whole number." },
        { "DeleteColumns", """{"Column":"A","Count":-2}""", "Count must be 1 or more, but it is -2." },

        // SetColumnWidth
        { "SetColumnWidth", """{"Width":10}""", "Missing required property 'Columns'." },
        { "SetColumnWidth", """{"Columns":"A"}""", "Missing required property 'Width'." },
        { "SetColumnWidth", """{"Column":"A","Width":10}""", "Unknown property 'Column'." },
        { "SetColumnWidth", """{"Columns":"A","Width":"wide"}""", "Property 'Width' must be a number." },
        { "SetColumnWidth", """{"Columns":"A","Width":null}""", "Property 'Width' must not be null." },
        { "SetColumnWidth", """{"Columns":"A","Width":0}""", "Width must be more than 0 and at most 255, but it is 0." },
        { "SetColumnWidth", """{"Columns":"A","Width":255.5}""", "Width must be more than 0 and at most 255, but it is 255.5." },
        { "SetColumnWidth", """{"Columns":"D:B","Width":10}""", "Columns 'D:B' must go from left to right, e.g. B:D." },
        { "SetColumnWidth", """{"Columns":"B:C:D","Width":10}""", "Columns 'B:C:D' is not a valid column or span, e.g. B or B:D." },
        { "SetColumnWidth", """{"Columns":"B2","Width":10}""", "Columns 'B2' is not a valid column or span, e.g. B or B:D." },
        { "SetColumnWidth", """{"Columns":"","Width":10}""", "Columns '' is not a valid column or span, e.g. B or B:D." },

        // AutoFitColumns
        { "AutoFitColumns", """{"Width":10}""", "Unknown property 'Width'. Allowed properties: Sheet, Columns." },
        { "AutoFitColumns", """{"Columns":1}""", "Property 'Columns' must be a string." },
        { "AutoFitColumns", """{"Columns":"1:2"}""", "Columns '1:2' is not a valid column or span, e.g. B or B:D." },
        { "AutoFitColumns", """{"Columns":"A:"}""", "Columns 'A:' is not a valid column or span, e.g. B or B:D." },
        { "AutoFitColumns", """{"Sheet":"Nope"}""", "Sheet 'Nope' was not found." },

        // GetCellValue
        { "GetCellValue", "{}", "Missing required property 'Cell'." },
        { "GetCellValue", """{"Range":"A1"}""", "Unknown property 'Range'." },
        { "GetCellValue", """{"Cell":null}""", "Property 'Cell' must not be null." },
        { "GetCellValue", """{"Cell":{"Row":1}}""", "Property 'Cell' must be a string." },
        { "GetCellValue", """{"Cell":"A0"}""", "Cell 'A0' is not a valid cell address" },

        // GetRangeValues
        { "GetRangeValues", "{}", "Missing required property 'Range'." },
        { "GetRangeValues", """{"Cell":"A1"}""", "Unknown property 'Cell'." },
        { "GetRangeValues", """{"Range":["A1","B2"]}""", "Property 'Range' must be a string." },
        { "GetRangeValues", """{"Range":"A1:"}""", "Range 'A1:' is not a valid range address, e.g. A1:D20." },
        { "GetRangeValues", """{"Range":"A:A"}""", "Range 'A:A' is not a valid range address, e.g. A1:D20." },
        { "GetRangeValues", """{"Range":"A1:B2:C3"}""", "Range 'A1:B2:C3' is not a valid range address, e.g. A1:D20." },
        { "GetRangeValues", """{"Range":"A1:Z100000"}""", "Range 'A1:Z100000' has more than 1,000,000 cells; read it in smaller parts." },

        // GetUsedRange
        { "GetUsedRange", """{"Range":"A1"}""", "Unknown property 'Range'. Allowed properties: Sheet." },
        { "GetUsedRange", """{"Sheet":1}""", "Property 'Sheet' must be a string." },
        { "GetUsedRange", """{"Sheet":"Nope"}""", "Sheet 'Nope' was not found." },

        // GetSheetNames
        { "GetSheetNames", """{"Sheet":"Sheet1"}""", "Unknown property 'Sheet'. This operation takes no arguments; use {}." },
    };

    /// <summary>
    /// Gets payloads that are not a JSON object, for every operation.
    /// </summary>
    public static TheoryData<string, string, string> BadPayloadCases
    {
        get
        {
            var payloads = new (string Arguments, string Message)[]
            {
                (string.Empty, "Arguments is empty; it must be a JSON object, e.g. {}."),
                ("   ", "Arguments is empty; it must be a JSON object, e.g. {}."),
                ("null", "Arguments must be a JSON object, not null."),
                ("[]", "Arguments must be a JSON object, not an array."),
                ("\"A1\"", "Arguments must be a JSON object, not a string."),
                ("42", "Arguments must be a JSON object, not a number."),
                ("true", "Arguments must be a JSON object, not a boolean."),
                ("{", "Arguments is not valid JSON:"),
                ("{\"Cell\":\"A1\",}", "Arguments is not valid JSON:"),
                ("{Cell:'A1'}", "Arguments is not valid JSON:"),
            };

            var data = new TheoryData<string, string, string>();
            foreach (var operation in CommandRunner.Operations)
            {
                foreach (var (arguments, message) in payloads)
                {
                    data.Add(operation.Name, arguments, message);
                }
            }

            return data;
        }
    }

    public static TheoryData<string> OperationNames() => [.. CommandRunner.Operations.Select(o => o.Name)];

    [Theory]
    [MemberData(nameof(ValidCases))]
    public void ValidArguments_Succeed(string operation, string arguments)
    {
        var ex = Record.Exception(() => Run(Cmd(operation, arguments)));

        Assert.Null(ex);
    }

    [Theory]
    [MemberData(nameof(InvalidCases))]
    [MemberData(nameof(BadPayloadCases))]
    public void InvalidArguments_FailWithAMessageNamingTheArgument(string operation, string arguments, string message)
    {
        var ex = Fail(operation, arguments);

        Assert.StartsWith($"Command #1 ({operation}): {message}", ex.Message, StringComparison.Ordinal);
        Assert.IsType<CommandArgumentException>(ex.InnerException);
    }

    [Fact]
    public void NullArguments_Fail()
    {
        var ex = Fail("GetSheetNames", null!);

        Assert.Equal("Command #1 (GetSheetNames): Arguments is empty; it must be a JSON object, e.g. {}.", ex.Message);
    }

    // --- Completeness: a new operation cannot ship without parameter tests ---------------------
    [Fact]
    public void Tables_OnlyNameRealOperations()
    {
        var names = CommandRunner.Operations.Select(o => o.Name).ToHashSet(StringComparer.Ordinal);

        Assert.All(ValidCases, row => Assert.Contains(row.Data.Item1, names));
        Assert.All(InvalidCases, row => Assert.Contains(row.Data.Item1, names));
    }

    [Fact]
    public void EveryOperation_HasAValidCase()
    {
        var covered = ValidCases.Select(row => row.Data.Item1).ToHashSet(StringComparer.Ordinal);

        Assert.All(CommandRunner.Operations, o => Assert.Contains(o.Name, covered));
    }

    [Fact]
    public void EveryOperation_HasAnUnknownPropertyCase()
    {
        var covered = InvalidCases
            .Where(row => row.Data.Item3.StartsWith("Unknown property", StringComparison.Ordinal))
            .Select(row => row.Data.Item1)
            .ToHashSet(StringComparer.Ordinal);

        Assert.All(CommandRunner.Operations, o => Assert.Contains(o.Name, covered));
    }

    [Fact]
    public void EveryRequiredProperty_HasAMissingPropertyCase()
    {
        var cases = InvalidCases.Select(row => (row.Data.Item1, row.Data.Item3)).ToHashSet();

        var missing =
            from operation in CommandRunner.Operations
            from property in CommandArgs.Options.GetTypeInfo(operation.ArgumentsType).Properties
            where property.IsRequired
            let expected = $"Missing required property '{property.Name}'."
            where !cases.Contains((operation.Name, expected))
            select $"{operation.Name}: {expected}";

        Assert.Empty(missing);
    }

    [Fact]
    public void EveryOptionalProperty_IsUsedInAValidCase()
    {
        var valid = ValidCases.Select(row => (Operation: row.Data.Item1, Arguments: row.Data.Item2)).ToList();

        var unused =
            from operation in CommandRunner.Operations
            from property in CommandArgs.Options.GetTypeInfo(operation.ArgumentsType).Properties
            where !property.IsRequired
            where !valid.Any(v => v.Operation == operation.Name && v.Arguments.Contains($"\"{property.Name}\"", StringComparison.Ordinal))
            select $"{operation.Name}.{property.Name}";

        Assert.Empty(unused);
    }

    [Theory]
    [MemberData(nameof(OperationNames))]
    public void EveryOperation_ParsesTheArgumentsTypeItDeclares(string operation)
    {
        // The catalog's ArgumentsType drives the completeness checks above, so it must be the
        // record the handler really parses. An unknown property lists that record's properties.
        var declared = CommandRunner.Operations.Single(o => o.Name == operation).ArgumentsType;
        var properties = CommandArgs.Options.GetTypeInfo(declared).Properties.Select(p => p.Name).ToList();
        var expected = properties.Count == 0
            ? "This operation takes no arguments; use {}."
            : $"Allowed properties: {string.Join(", ", properties)}.";

        var ex = Fail(operation, """{"__probe__":0}""");

        Assert.EndsWith($"Unknown property '__probe__'. {expected}", ex.Message, StringComparison.Ordinal);
    }
}
