using static OdcBatchExcel.Tests.Batch;

namespace OdcBatchExcel.Tests;

/// <summary>
/// The response must fit ODC's 5.5 MB payload limit; an oversized batch fails with a clear message.
/// </summary>
public class ResponseLimitTests
{
    private static readonly string LongText = new('x', 1000);

    [Fact]
    public void RealLimit_ReadJustUnderTheLimit_Succeeds()
    {
        // 5,000 cells x ~1 KB = ~5.0 MB of results. Identical strings share one entry in the file, so it stays small.
        var result = Run(TextColumn(5_000), Cmd("GetRangeValues", """{"Range":"A1:A5000"}"""));

        Assert.True(result.Results.Single().Value.Length > 5_000_000);
    }

    [Fact]
    public void RealLimit_ReadOverTheLimit_FailsNamingTheCommand()
    {
        // 6,000 cells x ~1 KB = ~6.0 MB of results: far below the 1,000,000-cell cap, but over 5.5 MB.
        var ex = Assert.Throws<InvalidOperationException>(
            () => Run(TextColumn(6_000), Cmd("GetSheetNames", "{}"), Cmd("GetRangeValues", """{"Range":"A1:A6000"}""")));

        Assert.StartsWith("Command #2 (GetRangeValues): The result is about 6,0", ex.Message, StringComparison.Ordinal);
        Assert.EndsWith("more than ODC's response limit of 5,500,000 bytes. Read a smaller range.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ResultsThatFitAlone_FailWhenTheyAddUpPastTheLimit()
    {
        // Each GetCellValue result is about 1,100 bytes, so two fit in 2,500 bytes and the third does not.
        var read = Cmd("GetCellValue", """{"Cell":"A1"}""");
        var commands = new List<Command> { read, read, Cmd("SetCellValue", """{"Cell":"B1","Value":"x"}"""), read };

        var ex = Assert.Throws<InvalidOperationException>(() => new OdcBatchExcel(2_500).Execute(TextColumn(1), commands));

        Assert.StartsWith("Command #4 (GetCellValue): The read results so far are about 3,", ex.Message, StringComparison.Ordinal);
        Assert.EndsWith("more than ODC's response limit of 2,500 bytes. Split the reads across several Execute calls.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FilePlusResults_FailWhenTogetherTheyExceedTheLimit()
    {
        var file = TextColumn(1);
        var fileAsBase64 = Convert.ToBase64String(Run(file).File).Length;
        var limit = fileAsBase64 + 500;

        // The file alone fits...
        var ok = Record.Exception(() => new OdcBatchExcel(limit).Execute(file, []));
        Assert.Null(ok);

        // ...but not together with a ~1,100-byte read result.
        var ex = Assert.Throws<InvalidOperationException>(
            () => new OdcBatchExcel(limit).Execute(file, [Cmd("GetCellValue", """{"Cell":"A1"}""")]));

        Assert.StartsWith("The response would be about ", ex.Message, StringComparison.Ordinal);
        Assert.Contains("(the file as Base64: ", ex.Message, StringComparison.Ordinal);
        Assert.EndsWith("Split the batch across several Execute calls, or make the workbook smaller.", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("", 0)]
    [InlineData("abc", 3)]
    [InlineData("日本語", 9)]
    [InlineData("a\"b\\", 6)]
    public void EncodedSize_CountsUtf8BytesAndJsonEscapes(string text, long expected)
    {
        Assert.Equal(expected, ResponseBudget.EncodedSize(text));
    }

    [Fact]
    public void Results_KeepNonAsciiTextUnescaped()
    {
        var result = Run(
            Cmd("SetCellValue", """{"Cell":"A1","Value":"日本語 <b>&'"}"""),
            Cmd("GetCellValue", """{"Cell":"A1"}"""));

        Assert.Equal("""{"Value":"日本語 <b>&'","Type":"Text"}""", result.Results.Single().Value);
    }

    // A sheet whose column A has the same 1,000-character text in rows 1..rows.
    private static byte[] TextColumn(int rows) => Workbook(w =>
    {
        var sheet = w.AddWorksheet("S");
        sheet.Range(1, 1, rows, 1).Value = LongText;
    });
}
