using System.Reflection;
using OutSystems.ExternalLibraries.SDK;

namespace OdcBatchExcel.Tests;

/// <summary>
/// Guards the ODC-facing contract: changing any of this breaks the apps that use the library.
/// </summary>
public class ContractTests
{
    [Fact]
    public void Interface_KeepsStableOdcLibraryName()
    {
        // The OSInterface name is the library's identity in ODC; renaming it creates a new library.
        var attribute = typeof(IOdcBatchExcel).GetCustomAttribute<OSInterfaceAttribute>();

        Assert.NotNull(attribute);
        Assert.Equal("OdcBatchExcel", attribute.Name);
    }

    [Fact]
    public void Interface_ExposesOnlyExecute()
    {
        var method = Assert.Single(typeof(IOdcBatchExcel).GetMethods());

        Assert.Equal("Execute", method.Name);
        Assert.NotNull(method.GetCustomAttribute<OSActionAttribute>());
        Assert.Equal(typeof(ExecuteResult), method.ReturnType);
        Assert.Equal(
            [(typeof(byte[]), "file"), (typeof(List<Command>), "commands")],
            method.GetParameters().Select(p => (p.ParameterType, p.Name!)));
        Assert.All(method.GetParameters(), p => Assert.NotNull(p.GetCustomAttribute<OSParameterAttribute>()));
    }

    [Theory]
    [InlineData(typeof(Command), new[] { "Operation:String", "Arguments:String" })]
    [InlineData(typeof(ExecuteResult), new[] { "File:Byte[]", "Results:List`1" })]
    [InlineData(typeof(CommandResult), new[] { "CommandIndex:Int32", "Operation:String", "Value:String" })]
    public void Structures_KeepTheirFields(Type type, string[] fields)
    {
        Assert.True(type.IsValueType);
        Assert.NotNull(type.GetCustomAttribute<OSStructureAttribute>());

        var properties = type.GetProperties();
        Assert.Equal(fields, properties.Select(p => $"{p.Name}:{p.PropertyType.Name}"));
        Assert.All(properties, p => Assert.NotNull(p.GetCustomAttribute<OSStructureFieldAttribute>()));
    }
}
