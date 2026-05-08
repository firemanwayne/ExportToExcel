using NPOI.XSSF.UserModel;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class HeaderBuilderTests
{
    [TestMethod]
    public void Build_WithDisplayAttribute_UsesDisplayName()
    {
        var workbook = new XSSFWorkbook();
        var builder = new HeaderBuilder<DisplayModel>(new HeaderStyle(), workbook);

        builder.Build("Sheet");

        var headers = builder.HeaderCells.Select(c => c.StringCellValue).ToList();
        CollectionAssert.Contains(headers, "First Name");
        CollectionAssert.Contains(headers, "Age");
        CollectionAssert.Contains(headers, "Score");
    }

    [TestMethod]
    public void Build_WithDisplayAttribute_CorrectColumnCount()
    {
        var workbook = new XSSFWorkbook();
        var builder = new HeaderBuilder<DisplayModel>(new HeaderStyle(), workbook);

        builder.Build("Sheet");

        Assert.AreEqual(3, builder.HeaderCells.Count());
    }

    [TestMethod]
    public void Build_WithSpreadSheetAttribute_UsesColumnName()
    {
        var workbook = new XSSFWorkbook();
        var builder = new HeaderBuilder<AttributeModel>(new HeaderStyle(), workbook);

        builder.Build("Sheet");

        var headers = builder.HeaderCells.Select(c => c.StringCellValue).ToList();
        CollectionAssert.Contains(headers, "Column A");
        CollectionAssert.Contains(headers, "Column B");
    }

    [TestMethod]
    public void Build_WithSpreadSheetAttribute_RespectsColumnOrder()
    {
        var workbook = new XSSFWorkbook();
        var builder = new HeaderBuilder<AttributeModel>(new HeaderStyle(), workbook);

        builder.Build("Sheet");

        var headers = builder.HeaderCells.OrderBy(c => c.ColumnIndex).Select(c => c.StringCellValue).ToList();
        Assert.AreEqual("Column A", headers[0]);
        Assert.AreEqual("Column B", headers[1]);
    }

    [TestMethod]
    public void Build_WithPlainModel_FallsBackToPropertyName()
    {
        var workbook = new XSSFWorkbook();
        var builder = new HeaderBuilder<PlainModel>(new HeaderStyle(), workbook);

        builder.Build("Sheet");

        var headers = builder.HeaderCells.Select(c => c.StringCellValue).ToList();
        CollectionAssert.Contains(headers, "Name");
        CollectionAssert.Contains(headers, "Value");
    }

    [TestMethod]
    public void ColumnProperties_CountMatchesDisplayModel()
    {
        var workbook = new XSSFWorkbook();
        var builder = new HeaderBuilder<DisplayModel>(new HeaderStyle(), workbook);

        builder.Build("Sheet");

        Assert.AreEqual(3, builder.ColumnProperties.Count());
    }

    [TestMethod]
    public void Build_CreatesSheetWithCorrectName()
    {
        var workbook = new XSSFWorkbook();
        var builder = new HeaderBuilder<DisplayModel>(new HeaderStyle(), workbook);

        var sheet = builder.Build("MySheet");

        Assert.AreEqual("MySheet", sheet.SheetName);
    }
}
