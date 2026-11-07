using ClosedXML.Excel;
using static OdcBatchExcel.Tests.Batch;

namespace OdcBatchExcel.Tests;

/// <summary>
/// Behavior of <see cref="OdcBatchExcel.Execute"/>: each operation's happy path, reads, batching and general errors.
/// </summary>
public class ExecuteTests
{
    // --- Input file ---------------------------------------------------------------------------
    [Fact]
    public void EmptyFile_StartsNewWorkbookWithSheet1()
    {
        using var workbook = Open(Run());

        Assert.Equal(["Sheet1"], workbook.Worksheets.Select(w => w.Name));
    }

    [Fact]
    public void NullFile_StartsNewWorkbook()
    {
        var result = new OdcBatchExcel().Execute(null!, []);

        using var workbook = Open(result);
        Assert.Equal("Sheet1", workbook.Worksheets.Single().Name);
    }

    [Fact]
    public void ExistingFile_IsChangedAndKeepsItsData()
    {
        var result = Run(DataWorkbook(), Cmd("SetCellValue", """{"Cell":"C1","Value":"c1"}"""));

        using var workbook = Open(result);
        var sheet = workbook.Worksheet("Data");
        Assert.Equal("a1", sheet.Cell("A1").GetText());
        Assert.Equal("c1", sheet.Cell("C1").GetText());
        Assert.Equal(["Data", "Other"], workbook.Worksheets.Select(w => w.Name));
    }

    [Fact]
    public void Output_AsksSpreadsheetAppsToRecalculateOnLoad()
    {
        using var workbook = Open(Run(Cmd("SetFormula", """{"Cell":"A1","Formula":"1+1"}""")));

        Assert.True(workbook.FullCalculationOnLoad);
    }

    [Fact]
    public void InvalidFile_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => Run([1, 2, 3]));

        Assert.StartsWith("File is not a valid .xlsx workbook", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void NullCommands_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => new OdcBatchExcel().Execute([], null!));
    }

    [Fact]
    public void NoCommands_ReturnsTheWorkbookAndNoResults()
    {
        var result = Run(DataWorkbook());

        Assert.Empty(result.Results);
        using var workbook = Open(result);
        Assert.Equal("b3", workbook.Worksheet("Data").Cell("B3").GetText());
    }

    // --- Sheet selection ----------------------------------------------------------------------
    [Theory]
    [InlineData("""{"Cell":"A1","Value":"x"}""", "Data")]
    [InlineData("""{"Sheet":"","Cell":"A1","Value":"x"}""", "Data")]
    [InlineData("""{"Sheet":"Other","Cell":"A1","Value":"x"}""", "Other")]
    [InlineData("""{"Sheet":"other","Cell":"A1","Value":"x"}""", "Other")]
    public void Sheet_DefaultsToTheFirstAndMatchesCaseInsensitively(string arguments, string expectedSheet)
    {
        using var workbook = Open(Run(DataWorkbook(), Cmd("SetCellValue", arguments)));

        Assert.Equal("x", workbook.Worksheet(expectedSheet).Cell("A1").GetText());
    }

    [Fact]
    public void Sheet_NotFound_ListsTheWorksheets()
    {
        var ex = Assert.Throws<InvalidOperationException>(
            () => Run(DataWorkbook(), Cmd("SetCellValue", """{"Sheet":"Nope","Cell":"A1","Value":"x"}""")));

        Assert.Equal("Command #1 (SetCellValue): Sheet 'Nope' was not found. Worksheets: 'Data', 'Other'.", ex.Message);
    }

    // --- Cell values and formulas -------------------------------------------------------------
    [Fact]
    public void SetCellValue_TextIsTheDefaultAndKeepsNumericLookingStrings()
    {
        using var workbook = Open(Run(Cmd("SetCellValue", """{"Cell":"B2","Value":"00123"}""")));

        var value = workbook.Worksheet(1).Cell("B2").Value;
        Assert.Equal(XLDataType.Text, value.Type);
        Assert.Equal("00123", value.GetText());
    }

    [Theory]
    [InlineData("Number", "1234.5", 1234.5)]
    [InlineData("number", "-1e3", -1000.0)]
    [InlineData("Number", "0", 0.0)]
    public void SetCellValue_Number(string valueType, string value, double expected)
    {
        using var workbook = Open(Run(Cmd("SetCellValue", $$"""{"Cell":"A1","Value":"{{value}}","ValueType":"{{valueType}}"}""")));

        Assert.Equal(expected, workbook.Worksheet(1).Cell("A1").Value.GetNumber());
    }

    [Theory]
    [InlineData("true", true)]
    [InlineData("False", false)]
    public void SetCellValue_Boolean(string value, bool expected)
    {
        using var workbook = Open(Run(Cmd("SetCellValue", $$"""{"Cell":"A1","Value":"{{value}}","ValueType":"Boolean"}""")));

        Assert.Equal(expected, workbook.Worksheet(1).Cell("A1").Value.GetBoolean());
    }

    [Theory]
    [InlineData("2026-11-07", "2026-11-07T00:00:00")]
    [InlineData("2026-11-07T09:30:15", "2026-11-07T09:30:15")]
    [InlineData("2026-11-07T09:30:00Z", "2026-11-07T09:30:00")]
    [InlineData("2026-11-07T09:30:00+09:00", "2026-11-07T00:30:00")]
    public void SetCellValue_DateTime_ConvertsOffsetsToUtc(string value, string expected)
    {
        using var workbook = Open(Run(Cmd("SetCellValue", $$"""{"Cell":"A1","Value":"{{value}}","ValueType":"DateTime"}""")));

        var cell = workbook.Worksheet(1).Cell("A1").Value;
        Assert.Equal(XLDataType.DateTime, cell.Type);
        Assert.Equal(DateTime.Parse(expected, System.Globalization.CultureInfo.InvariantCulture), cell.GetDateTime());
    }

    [Theory]
    [InlineData("Text")]
    [InlineData("Number")]
    [InlineData("Boolean")]
    [InlineData("DateTime")]
    public void SetCellValue_EmptyValueClearsTheCell(string valueType)
    {
        var result = Run(DataWorkbook(), Cmd("SetCellValue", $$"""{"Cell":"A1","Value":"","ValueType":"{{valueType}}"}"""));

        using var workbook = Open(result);
        Assert.True(workbook.Worksheet("Data").Cell("A1").Value.IsBlank);
    }

    [Theory]
    [InlineData("SUM(A1:A2)")]
    [InlineData("=SUM(A1:A2)")]
    public void SetFormula_StoresTheFormulaWithOrWithoutEquals(string formula)
    {
        var result = Run(
            Cmd("SetCellValue", """{"Cell":"A1","Value":"1","ValueType":"Number"}"""),
            Cmd("SetCellValue", """{"Cell":"A2","Value":"2","ValueType":"Number"}"""),
            Cmd("SetFormula", $$"""{"Cell":"A3","Formula":"{{formula}}"}"""));

        using var workbook = Open(result);
        var cell = workbook.Worksheet(1).Cell("A3");
        Assert.Equal("SUM(A1:A2)", cell.FormulaA1);
        Assert.Equal(3.0, cell.Value.GetNumber());
    }

    [Fact]
    public void SetRangeValues_WritesABlockWithColumnTypes()
    {
        var result = Run(Cmd("SetRangeValues", """
            {
              "StartCell": "B2",
              "Rows": [["Apple", "1.5", "true", "2026-01-31"], ["Pear", "", "false", "2026-02-01"]],
              "ColumnTypes": ["Text", "Number", "Boolean", "DateTime"]
            }
            """));

        using var workbook = Open(result);
        var sheet = workbook.Worksheet(1);
        Assert.Equal("Apple", sheet.Cell("B2").GetText());
        Assert.Equal(1.5, sheet.Cell("C2").Value.GetNumber());
        Assert.True(sheet.Cell("D2").Value.GetBoolean());
        Assert.Equal(new DateTime(2026, 1, 31, 0, 0, 0, DateTimeKind.Unspecified), sheet.Cell("E2").Value.GetDateTime());
        Assert.Equal("Pear", sheet.Cell("B3").GetText());
        Assert.True(sheet.Cell("C3").Value.IsBlank);
        Assert.False(sheet.Cell("D3").Value.GetBoolean());
        Assert.Equal("B2:E3", sheet.RangeUsed()!.RangeAddress.ToStringRelative());
    }

    [Fact]
    public void SetRangeValues_WithoutColumnTypes_WritesTextAndAllowsRowsOfDifferentLengths()
    {
        var result = Run(Cmd("SetRangeValues", """{"StartCell":"A1","Rows":[["1","2","3"],["4"]]}"""));

        using var workbook = Open(result);
        var sheet = workbook.Worksheet(1);
        Assert.Equal("3", sheet.Cell("C1").Value.GetText());
        Assert.Equal("4", sheet.Cell("A2").Value.GetText());
        Assert.True(sheet.Cell("B2").Value.IsBlank);
    }

    [Fact]
    public void SetRangeValues_BadValue_LeavesTheSheetUnchanged()
    {
        // The whole batch fails, so no file comes back; this checks the command itself is atomic too.
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");

        Assert.Throws<Commands.CommandArgumentException>(() => Commands.CommandRunner.Run(
            workbook,
            "SetRangeValues",
            """{"StartCell":"A1","Rows":[["1"],["x"]],"ColumnTypes":["Number"]}"""));
        Assert.True(sheet.Cell("A1").Value.IsBlank);
    }

    // --- Rows and columns ---------------------------------------------------------------------
    [Theory]
    [InlineData("""{"Row":2}""", new[] { "a1", "", "a2", "a3" })]
    [InlineData("""{"Row":1,"Count":2}""", new[] { "", "", "a1", "a2", "a3" })]
    public void InsertRows_ShiftsDataDown(string arguments, string[] expectedColumnA)
    {
        using var workbook = Open(Run(DataWorkbook(), Cmd("InsertRows", arguments)));

        Assert.Equal(expectedColumnA, ColumnA(workbook, expectedColumnA.Length));
    }

    [Theory]
    [InlineData("""{"Row":2}""", new[] { "a1", "a3", "" })]
    [InlineData("""{"Row":1,"Count":2}""", new[] { "a3", "", "" })]
    public void DeleteRows_ShiftsDataUp(string arguments, string[] expectedColumnA)
    {
        using var workbook = Open(Run(DataWorkbook(), Cmd("DeleteRows", arguments)));

        Assert.Equal(expectedColumnA, ColumnA(workbook, expectedColumnA.Length));
    }

    [Fact]
    public void InsertColumns_ShiftsDataRight()
    {
        using var workbook = Open(Run(DataWorkbook(), Cmd("InsertColumns", """{"Column":"B","Count":2}""")));

        var sheet = workbook.Worksheet("Data");
        Assert.Equal("a1", sheet.Cell("A1").GetText());
        Assert.True(sheet.Cell("B1").Value.IsBlank);
        Assert.True(sheet.Cell("C1").Value.IsBlank);
        Assert.Equal("b1", sheet.Cell("D1").GetText());
    }

    [Fact]
    public void DeleteColumns_ShiftsDataLeft()
    {
        using var workbook = Open(Run(DataWorkbook(), Cmd("DeleteColumns", """{"Column":"a"}""")));

        var sheet = workbook.Worksheet("Data");
        Assert.Equal("b1", sheet.Cell("A1").GetText());
        Assert.True(sheet.Cell("B1").Value.IsBlank);
    }

    [Fact]
    public void InsertRows_MovesFormulaReferences()
    {
        var result = Run(
            DataWorkbook(),
            Cmd("SetFormula", """{"Cell":"C1","Formula":"A3"}"""),
            Cmd("InsertRows", """{"Row":2}"""),
            Cmd("GetCellValue", """{"Cell":"C1"}"""));

        Assert.Equal("""{"Value":"a3","Type":"Text"}""", result.Results.Single().Value);
        using var workbook = Open(result);
        Assert.Equal("A4", workbook.Worksheet("Data").Cell("C1").FormulaA1);
    }

    [Theory]
    [InlineData("B", 30.0, new[] { "B" })]
    [InlineData("B:D", 12.5, new[] { "B", "C", "D" })]
    public void SetColumnWidth_SetsEveryColumnInTheSpan(string columns, double width, string[] expected)
    {
        using var workbook = Open(Run(Cmd("SetColumnWidth", $$"""{"Columns":"{{columns}}","Width":{{width}}}""")));

        var sheet = workbook.Worksheet(1);
        Assert.All(expected, c => Assert.Equal(width, sheet.Column(c).Width));
        Assert.NotEqual(width, sheet.Column("E").Width);
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("""{"Columns":"A"}""")]
    public void AutoFitColumns_WidensAColumnWithLongText(string arguments)
    {
        var file = Workbook(w => w.AddWorksheet("S").Cell("A1").Value = new string('W', 60));
        using var before = new XLWorkbook(new MemoryStream(file));
        var defaultWidth = before.Worksheet("S").Column("A").Width;

        using var workbook = Open(Run(file, Cmd("AutoFitColumns", arguments)));

        Assert.True(workbook.Worksheet("S").Column("A").Width > defaultWidth * 3);
    }

    [Fact]
    public void AutoFitColumns_EmptySheet_DoesNothing()
    {
        var result = Run(Cmd("AutoFitColumns", "{}"));

        Assert.Empty(result.Results);
    }

    // --- Reads --------------------------------------------------------------------------------
    [Theory]
    [InlineData("""{"Cell":"A1","Value":"hello"}""", """{"Value":"hello","Type":"Text"}""")]
    [InlineData("""{"Cell":"A1","Value":"1234.5","ValueType":"Number"}""", """{"Value":"1234.5","Type":"Number"}""")]
    [InlineData("""{"Cell":"A1","Value":"true","ValueType":"Boolean"}""", """{"Value":"true","Type":"Boolean"}""")]
    [InlineData("""{"Cell":"A1","Value":"2026-11-07T09:30:00","ValueType":"DateTime"}""", """{"Value":"2026-11-07T09:30:00","Type":"DateTime"}""")]
    [InlineData("""{"Cell":"A1","Value":""}""", """{"Value":"","Type":"Blank"}""")]
    public void GetCellValue_ReturnsValueAndType(string setArguments, string expected)
    {
        var result = Run(Cmd("SetCellValue", setArguments), Cmd("GetCellValue", """{"Cell":"A1"}"""));

        Assert.Equal(expected, result.Results.Single().Value);
    }

    [Theory]
    [InlineData("1/0", """{"Value":"#DIV/0!","Type":"Error"}""")]
    [InlineData("NOSUCHFUNCTION(1)", """{"Value":"#NAME?","Type":"Error"}""")]
    [InlineData("SQRT(-1)", """{"Value":"#NUM!","Type":"Error"}""")]
    [InlineData("NA()", """{"Value":"#N/A","Type":"Error"}""")]
    [InlineData("\\u0022a\\u0022+1", """{"Value":"#VALUE!","Type":"Error"}""")]
    [InlineData("\\u0022a\\u0022&\\u0022b\\u0022", """{"Value":"ab","Type":"Text"}""")]
    [InlineData("2>1", """{"Value":"true","Type":"Boolean"}""")]
    public void GetCellValue_ReturnsTheCalculatedValueOfAFormula(string formula, string expected)
    {
        var result = Run(Cmd("SetFormula", $$"""{"Cell":"A1","Formula":"{{formula}}"}"""), Cmd("GetCellValue", """{"Cell":"A1"}"""));

        Assert.Equal(expected, result.Results.Single().Value);
    }

    [Fact]
    public void DeleteRows_ReferenceToADeletedRow_MovesToTheRowAbove_UnlikeExcel()
    {
        // Known ClosedXML limitation (documented in docs/design.md): Excel turns a reference to a
        // deleted row into #REF!, but ClosedXML points it at the row above. This test pins the
        // current behavior so a ClosedXML upgrade that changes it is noticed.
        var result = Run(
            DataWorkbook(),
            Cmd("SetFormula", """{"Cell":"C1","Formula":"A3"}"""),
            Cmd("DeleteRows", """{"Row":3}"""),
            Cmd("GetCellValue", """{"Cell":"C1"}"""));

        Assert.Equal("""{"Value":"a2","Type":"Text"}""", result.Results.Single().Value);
        using var workbook = Open(result);
        Assert.Equal("A2", workbook.Worksheet("Data").Cell("C1").FormulaA1);
    }

    [Fact]
    public void GetCellValue_FormulaClosedXmlCannotCalculate_NamesTheCell()
    {
        // ClosedXML does not implement every Excel feature, e.g. the range intersection operator (a space).
        var ex = Assert.Throws<InvalidOperationException>(() => Run(
            Cmd("SetFormula", """{"Cell":"A1","Formula":"SUM(B1:B2 C1:C2)"}"""),
            Cmd("GetCellValue", """{"Cell":"A1"}""")));

        Assert.StartsWith("Command #2 (GetCellValue): Could not calculate the formula in A1: ", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void GetCellValue_TimeSpan_UsesInvariantFormat()
    {
        var file = Workbook(w => w.AddWorksheet("S").Cell("A1").Value = new TimeSpan(1, 2, 30, 15));

        var result = Run(file, Cmd("GetCellValue", """{"Cell":"A1"}"""));

        Assert.Equal("""{"Value":"1.02:30:15","Type":"TimeSpan"}""", result.Results.Single().Value);
    }

    [Theory]
    [InlineData("A1:B3", """{"Rows":[["a1","b1"],["a2","b2"],["a3","b3"]]}""")]
    [InlineData("B2", """{"Rows":[["b2"]]}""")]
    [InlineData("B3:C4", """{"Rows":[["b3",""],["",""]]}""")]
    public void GetRangeValues_ReturnsRowsOfStrings(string range, string expected)
    {
        var result = Run(DataWorkbook(), Cmd("GetRangeValues", $$"""{"Range":"{{range}}"}"""));

        Assert.Equal(expected, result.Results.Single().Value);
    }

    [Fact]
    public void GetUsedRange_ReturnsAddressAndSize()
    {
        var result = Run(
            Cmd("SetCellValue", """{"Cell":"B3","Value":"x"}"""),
            Cmd("SetCellValue", """{"Cell":"D20","Value":"y"}"""),
            Cmd("GetUsedRange", "{}"));

        Assert.Equal("""{"Address":"B3:D20","FirstRow":3,"FirstColumn":2,"RowCount":18,"ColumnCount":3}""", result.Results.Single().Value);
    }

    [Fact]
    public void GetUsedRange_EmptySheet_ReturnsZeros()
    {
        var result = Run(Cmd("GetUsedRange", """{"Sheet":"Sheet1"}"""));

        Assert.Equal("""{"Address":"","FirstRow":0,"FirstColumn":0,"RowCount":0,"ColumnCount":0}""", result.Results.Single().Value);
    }

    [Fact]
    public void GetSheetNames_ReturnsNamesInWorkbookOrder()
    {
        var result = Run(DataWorkbook(), Cmd("GetSheetNames", "{}"));

        Assert.Equal("""{"Names":["Data","Other"]}""", result.Results.Single().Value);
    }

    [Fact]
    public void Reads_ReturnJsonThatRoundTripsNonAsciiTextAndQuotes()
    {
        var result = Run(
            Cmd("SetCellValue", """{"Cell":"A1","Value":"日本語 \"quoted\""}"""),
            Cmd("GetCellValue", """{"Cell":"A1"}"""));

        using var json = System.Text.Json.JsonDocument.Parse(result.Results.Single().Value);
        Assert.Equal("日本語 \"quoted\"", json.RootElement.GetProperty("Value").GetString());
    }

    // --- Batching -----------------------------------------------------------------------------
    [Fact]
    public void Batch_RunsInOrderAndReturnsOneResultPerRead()
    {
        var result = Run(
            Cmd("SetRangeValues", """{"StartCell":"A1","Rows":[["2","3"]],"ColumnTypes":["Number","Number"]}"""),
            Cmd("SetFormula", """{"Cell":"C1","Formula":"A1*B1"}"""),
            Cmd("GetCellValue", """{"Cell":"C1"}"""),
            Cmd("SetCellValue", """{"Cell":"A1","Value":"10","ValueType":"Number"}"""),
            Cmd("getcellvalue", """{"Cell":"C1"}"""),
            Cmd("InsertColumns", """{"Column":"A"}"""),
            Cmd("GetRangeValues", """{"Range":"A1:D1"}"""));

        Assert.Equal(
            [
                (3, "GetCellValue", """{"Value":"6","Type":"Number"}"""),
                (5, "GetCellValue", """{"Value":"30","Type":"Number"}"""),
                (7, "GetRangeValues", """{"Rows":[["","10","3","30"]]}"""),
            ],
            result.Results.Select(r => (r.CommandIndex, r.Operation, r.Value)));

        using var workbook = Open(result);
        Assert.Equal("B1*C1", workbook.Worksheet(1).Cell("D1").FormulaA1);
    }

    // --- General errors -----------------------------------------------------------------------
    [Fact]
    public void Error_NamesTheFailingCommandByIndexAndOperation()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => Run(
            Cmd("SetCellValue", """{"Cell":"A1","Value":"1","ValueType":"Number"}"""),
            Cmd("SetCellValue", """{"Cell":"A2","Value":"1,5","ValueType":"Number"}""")));

        Assert.Equal(
            "Command #2 (SetCellValue): Value '1,5' is not a valid Number. Use the invariant format, e.g. 1234.5.",
            ex.Message);
        Assert.IsType<Commands.CommandArgumentException>(ex.InnerException);
    }

    [Fact]
    public void Error_UnknownOperation_ListsTheSupportedOperations()
    {
        var ex = Fail("SetValue", "{}");

        Assert.StartsWith("Command #1 (SetValue): Unknown operation 'SetValue'. Supported: SetCellValue, SetFormula,", ex.Message, StringComparison.Ordinal);
        Assert.EndsWith("GetSheetNames.", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData(null)]
    public void Error_EmptyOperation(string? operation)
    {
        var ex = Fail(operation!, "{}");

        Assert.EndsWith("Operation is empty.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Error_DefaultCommandStruct_IsRejected()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => Run(default(Command)));

        Assert.Equal("Command #1 (): Operation is empty.", ex.Message);
    }

    private static string[] ColumnA(XLWorkbook workbook, int rows)
    {
        var sheet = workbook.Worksheet("Data");
        return [.. Enumerable.Range(1, rows).Select(r => sheet.Cell(r, 1).Value.IsBlank ? string.Empty : sheet.Cell(r, 1).GetText())];
    }
}
