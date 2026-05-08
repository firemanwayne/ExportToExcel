using Microsoft.AspNetCore.Components.Forms;
using Simple.ExportToExcel.Models.Concrete;
using Simple.ExportToExcel.Styles;

namespace Simple.ExportToExcel.UnitTests;

// Minimal IBrowserFile fake for SpreadSheet constructor tests
file sealed class FakeBrowserFile : IBrowserFile
{
    public string Name { get; init; } = "upload.xlsx";
    public DateTimeOffset LastModified { get; init; } = DateTimeOffset.UtcNow;
    public long Size { get; init; } = 4096;
    public string ContentType { get; init; } = ExcelConstants.Excel;
    public Stream OpenReadStream(long maxAllowedSize = 512000, CancellationToken cancellationToken = default)
        => Stream.Null;
}

[TestClass]
public class SpreadSheetTests
{
    [TestMethod]
    public void DefaultConstructor_RowCountIsZero()
    {
        var sheet = new SpreadSheet();

        Assert.AreEqual(0, sheet.RowCount);
    }

    [TestMethod]
    public void DefaultConstructor_DataRowsIsNotNull()
    {
        var sheet = new SpreadSheet();

        Assert.IsNotNull(sheet.DataRows);
    }

    [TestMethod]
    public void BrowserFileConstructor_SetsFileName()
    {
        var file = new FakeBrowserFile { Name = "report.xlsx" };

        var sheet = new SpreadSheet(file);

        Assert.AreEqual("report.xlsx", sheet.FileName);
    }

    [TestMethod]
    public void BrowserFileConstructor_SetsSize()
    {
        var file = new FakeBrowserFile { Size = 8192 };

        var sheet = new SpreadSheet(file);

        Assert.AreEqual(8192, sheet.Size);
    }

    [TestMethod]
    public void BrowserFileConstructor_SetsContentType()
    {
        var file = new FakeBrowserFile { ContentType = ExcelConstants.Excel };

        var sheet = new SpreadSheet(file);

        Assert.AreEqual(ExcelConstants.Excel, sheet.ContentType);
    }

    [TestMethod]
    public void AddRow_Object_IncrementsRowCount()
    {
        var sheet = new SpreadSheet();

        sheet.AddRow("a,b,c");

        Assert.AreEqual(1, sheet.RowCount);
    }

    [TestMethod]
    public void AddRow_SheetRow_IncrementsRowCount()
    {
        var sheet = new SpreadSheet();
        var row = new SheetRow(0, "hello");

        sheet.AddRow(row);

        Assert.AreEqual(1, sheet.RowCount);
    }

    [TestMethod]
    public void AddRow_MultipleTimes_RowCountAccumulates()
    {
        var sheet = new SpreadSheet();

        sheet.AddRow("a,b").AddRow("c,d").AddRow("e,f");

        Assert.AreEqual(3, sheet.RowCount);
    }

    [TestMethod]
    public void RowCount_MatchesDataRowsCount()
    {
        var sheet = new SpreadSheet();
        sheet.AddRow("x").AddRow("y");

        Assert.AreEqual(sheet.RowCount, sheet.DataRows.Count);
    }

    [TestMethod]
    public void AddRow_Object_UsesCurrentRowCountAsIndex()
    {
        var sheet = new SpreadSheet();
        sheet.AddRow("first");   // index 0
        sheet.AddRow("second");  // index 1

        var rows = sheet.DataRows.ToList();
        Assert.AreEqual(0, rows[0].Index);
        Assert.AreEqual(1, rows[1].Index);
    }

    [TestMethod]
    public void AddRow_PreservesInsertionOrder()
    {
        var sheet = new SpreadSheet();
        var r0 = new SheetRow(0, "zero");
        var r1 = new SheetRow(1, "one");
        var r2 = new SheetRow(2, "two");

        sheet.AddRow(r0).AddRow(r1).AddRow(r2);

        var rows = sheet.DataRows.ToList();
        Assert.AreSame(r0, rows[0]);
        Assert.AreSame(r1, rows[1]);
        Assert.AreSame(r2, rows[2]);
    }

    [TestMethod]
    public void AddRow_ReturnsSpreadSheetForChaining()
    {
        var sheet = new SpreadSheet();

        var result = sheet.AddRow("x");

        Assert.AreSame(sheet, result);
    }

    [TestMethod]
    public void AddRow_SheetRow_ReturnsSpreadSheetForChaining()
    {
        var sheet = new SpreadSheet();
        var row = new SheetRow(0, "data");

        var result = sheet.AddRow(row);

        Assert.AreSame(sheet, result);
    }

    [TestMethod]
    public void DataRows_ContainsAddedRows()
    {
        var sheet = new SpreadSheet();
        var row = new SheetRow(0, "data");

        sheet.AddRow(row);

        CollectionAssert.Contains(sheet.DataRows.ToList(), row);
    }
}

[TestClass]
public class SheetRowTests
{
    [TestMethod]
    public void Constructor_SetsIndex()
    {
        var row = new SheetRow(3, "a,b,c");

        Assert.AreEqual(3, row.Index);
    }

    [TestMethod]
    public void Constructor_WithCommaSeparatedString_CreatesColumns()
    {
        var row = new SheetRow(0, "a,b,c");

        Assert.IsTrue(row.ColumnCount > 0);
    }

    [TestMethod]
    public void Constructor_WithNonStringValue_CreatesNoColumns()
    {
        var row = new SheetRow(0, 42);

        Assert.AreEqual(0, row.ColumnCount);
    }

    [TestMethod]
    public void Columns_IsNotNull()
    {
        var row = new SheetRow(0, "x");

        Assert.IsNotNull(row.Columns);
    }
}

[TestClass]
public class SheetColumnTests
{
    [TestMethod]
    public void Constructor_SetsRowIndex()
    {
        var col = new SheetColumn(2, 1, "value");

        Assert.AreEqual(2, col.RowIndex);
    }

    [TestMethod]
    public void Constructor_SetsColumnIndex()
    {
        var col = new SheetColumn(0, 5, "value");

        Assert.AreEqual(5, col.ColumnIndex);
    }

    [TestMethod]
    public void Constructor_SetsValue()
    {
        var col = new SheetColumn(0, 0, "hello");

        Assert.AreEqual("hello", col.Value);
    }

    [TestMethod]
    public void Constructor_AcceptsNullValue()
    {
        var col = new SheetColumn(0, 0, null);

        Assert.IsNull(col.Value);
    }
}

[TestClass]
public class ConditionalStyleTests
{
    static readonly ExcelColors Red  = ExcelColors.ColorCollection.Values.First(c => c.Name == "Red");
    static readonly ExcelColors Blue = ExcelColors.ColorCollection.Values.First(c => c.Name == "Blue");

    [TestMethod]
    public void Evaluate_ConditionTrue_ReturnsTrueColor()
    {
        var style = new ConditionalStyle(_ => true, "any", Red, Blue);

        var result = style.Evaluate();

        Assert.AreSame(Red, result);
    }

    [TestMethod]
    public void Evaluate_ConditionFalse_ReturnsFalseColor()
    {
        var style = new ConditionalStyle(_ => false, "any", Red, Blue);

        var result = style.Evaluate();

        Assert.AreSame(Blue, result);
    }

    [TestMethod]
    public void Evaluate_PassesValueToCondition()
    {
        string captured = null;
        var style = new ConditionalStyle(v => { captured = v; return true; }, "test-value", Red, Blue);

        style.Evaluate();

        Assert.AreEqual("test-value", captured);
    }

    [TestMethod]
    public void Constructor_SetsAllProperties()
    {
        Func<string, bool> condition = _ => true;
        var style = new ConditionalStyle(condition, "val", Red, Blue);

        Assert.AreSame(condition, style.Condition);
        Assert.AreEqual("val", style.Value);
        Assert.AreSame(Red, style.TrueColor);
        Assert.AreSame(Blue, style.FalseColor);
    }
}
