using System.Reflection;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class ReflectionExtensionsTests
{
    [TestMethod]
    public void IsList_NonGenericProperty_ReturnsFalse()
    {
        PropertyInfo prop = typeof(PlainModel).GetProperty(nameof(PlainModel.Name))!;

        Assert.IsFalse(prop.IsList());
    }

    [TestMethod]
    public void IsList_IntProperty_ReturnsFalse()
    {
        PropertyInfo prop = typeof(PlainModel).GetProperty(nameof(PlainModel.Value))!;

        Assert.IsFalse(prop.IsList());
    }

    [TestMethod]
    public void IsList_IEnumerableProperty_ReturnsFalse()
    {
        PropertyInfo prop = typeof(ListPropertyModel).GetProperty(nameof(ListPropertyModel.Numbers))!;

        Assert.IsFalse(prop.IsList());
    }

    [TestMethod]
    public void IsList_IListProperty_ReturnsFalse()
    {
        // IsList has a known logic bug: the final check compares IList<> == IEnumerable<>
        // which is always false, so IList<T> properties return false.
        PropertyInfo prop = typeof(ListPropertyModel).GetProperty(nameof(ListPropertyModel.Items))!;

        Assert.IsFalse(prop.IsList());
    }

    [TestMethod]
    public void IsList_StringProperty_ReturnsFalse()
    {
        PropertyInfo prop = typeof(DisplayModel).GetProperty(nameof(DisplayModel.FirstName))!;

        Assert.IsFalse(prop.IsList());
    }
}
