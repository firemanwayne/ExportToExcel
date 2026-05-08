using Simple.ExportToExcel.Attributes;
using System.Reflection;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class SpreadSheetAttributeTests
{
    [TestMethod]
    public void Columns_AreOrderedByColumnIndex()
    {
        var attr = typeof(AttributeModel).GetCustomAttribute<SpreadSheetAttribute>()!;

        var columns = attr.Columns.ToList();

        Assert.AreEqual(0, columns[0].ColumnIndex);
        Assert.AreEqual(1, columns[1].ColumnIndex);
    }

    [TestMethod]
    public void Columns_CountMatchesDecoratedProperties()
    {
        var attr = typeof(AttributeModel).GetCustomAttribute<SpreadSheetAttribute>()!;

        Assert.AreEqual(2, attr.Columns.Count());
    }

    [TestMethod]
    public void Name_IsSetFromConstructor()
    {
        var attr = typeof(AttributeModel).GetCustomAttribute<SpreadSheetAttribute>()!;

        Assert.AreEqual("AttributeModel", attr.Name);
    }

    [TestMethod]
    public void Type_IsSetFromConstructor()
    {
        var attr = typeof(AttributeModel).GetCustomAttribute<SpreadSheetAttribute>()!;

        Assert.AreEqual(typeof(AttributeModel), attr.Type);
    }

    [TestMethod]
    public void Columns_EachHasPropertySet()
    {
        var attr = typeof(AttributeModel).GetCustomAttribute<SpreadSheetAttribute>()!;

        foreach (var col in attr.Columns)
        {
            Assert.IsNotNull(col.Property);
        }
    }
}

[TestClass]
public class SpreadSheetColumnAttributeTests
{
    [TestMethod]
    public void Name_IsSetFromConstructor()
    {
        var prop = typeof(AttributeModel).GetProperty(nameof(AttributeModel.A))!;
        var attr = prop.GetCustomAttribute<SpreadSheetColumnAttribute>()!;

        Assert.AreEqual("Column A", attr.Name);
    }

    [TestMethod]
    public void ColumnIndex_IsSetFromConstructor()
    {
        var prop = typeof(AttributeModel).GetProperty(nameof(AttributeModel.A))!;
        var attr = prop.GetCustomAttribute<SpreadSheetColumnAttribute>()!;

        Assert.AreEqual(0, attr.ColumnIndex);
    }

    [TestMethod]
    public void Property_IsSetAfterSpreadSheetAttributeInitialization()
    {
        // SpreadSheetAttribute calls SetProperty during its constructor
        var attr = typeof(AttributeModel).GetCustomAttribute<SpreadSheetAttribute>()!;
        var colA = attr.Columns.First(c => c.ColumnIndex == 0);

        Assert.IsNotNull(colA.Property);
        Assert.AreEqual(nameof(AttributeModel.A), colA.Property.Name);
    }

    [TestMethod]
    public void ColumnB_HasCorrectName()
    {
        var prop = typeof(AttributeModel).GetProperty(nameof(AttributeModel.B))!;
        var attr = prop.GetCustomAttribute<SpreadSheetColumnAttribute>()!;

        Assert.AreEqual("Column B", attr.Name);
        Assert.AreEqual(1, attr.ColumnIndex);
    }
}
