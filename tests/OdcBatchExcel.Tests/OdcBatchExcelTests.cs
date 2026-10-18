using System.Reflection;
using OutSystems.ExternalLibraries.SDK;

namespace OdcBatchExcel.Tests;

public class OdcBatchExcelTests
{
    private readonly OdcBatchExcel sut = new OdcBatchExcel();

    [Theory]
    [InlineData("World", "Hello, World!")]
    [InlineData("  ODC  ", "Hello, ODC!")]
    public void Greet_ReturnsGreeting(string name, string expected)
    {
        Assert.Equal(expected, this.sut.Greet(name));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Greet_RejectsBlankName(string name)
    {
        Assert.Throws<ArgumentException>(() => this.sut.Greet(name));
    }

    [Fact]
    public void Greet_RejectsNullName()
    {
        Assert.Throws<ArgumentNullException>(() => this.sut.Greet(null!));
    }

    [Fact]
    public void Interface_KeepsStableOdcLibraryName()
    {
        // The OSInterface name is the library's identity in ODC; renaming it creates a new library.
        var attribute = typeof(IOdcBatchExcel).GetCustomAttribute<OSInterfaceAttribute>();

        Assert.NotNull(attribute);
        Assert.Equal("OdcBatchExcel", attribute.Name);
    }
}
