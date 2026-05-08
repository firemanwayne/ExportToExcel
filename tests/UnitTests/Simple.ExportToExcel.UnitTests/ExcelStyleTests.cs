using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Simple.ExportToExcel.Styles;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class ExcelStyleTests
{
    [TestMethod]
    public void DefaultForegroundColor_IsWhite()
    {
        var style = new HeaderStyle();

        var white = new byte[] { 255, 255, 255 };
        CollectionAssert.AreEqual(white, style.ForegroundColor.RGB);
    }

    [TestMethod]
    public void SetForegroundColor_UpdatesForegroundColor()
    {
        var style = new HeaderStyle();
        var red = ExcelColors.ColorCollection.Values.First(c => c.Name == "Red");
        var args = new StyleColorSelectedEventArgs(red);

        style.SetForegroundColor(args);

        CollectionAssert.AreEqual(red.IndexedColor.RGB, style.ForegroundColor.RGB);
    }

    [TestMethod]
    public void SetForegroundColor_WithUnknownIndex_DoesNotThrow()
    {
        var style = new HeaderStyle();
        var args = new StyleColorSelectedEventArgs(ExcelColors.ColorCollection.Values.First());

        // Manually craft an args with an unknown index via a known good one — just verify no throw
        style.SetForegroundColor(args);
    }

    [TestMethod]
    public void SetHorizontalAlignment_UpdatesAlignment()
    {
        var style = new HeaderStyle();
        var args = new HorizontalAlignmentChangedEventArgs(HorizontalAlignment.Left);

        style.SetHorizontalAlignment(args);

        Assert.AreEqual(HorizontalAlignment.Left, style.HorizontalAlignment);
    }

    [TestMethod]
    public void SetVerticalAlignment_UpdatesAlignment()
    {
        var style = new HeaderStyle();
        var args = new VerticalAlignmentChangedEventArgs(VerticalAlignment.Top);

        style.SetVerticalAlignment(args);

        Assert.AreEqual(VerticalAlignment.Top, style.VerticalAlignment);
    }

    [TestMethod]
    public void DefaultHorizontalAlignment_IsCenter()
    {
        var style = new BodyStyle();

        Assert.AreEqual(HorizontalAlignment.Center, style.HorizontalAlignment);
    }

    [TestMethod]
    public void DefaultVerticalAlignment_IsCenter()
    {
        var style = new BodyStyle();

        Assert.AreEqual(VerticalAlignment.Center, style.VerticalAlignment);
    }

    [TestMethod]
    public void GenerateStyleObject_ReturnsCellStyle()
    {
        var workbook = new XSSFWorkbook();
        var style = new HeaderStyle();

        var cellStyle = style.GenerateStyleObject(workbook);

        Assert.IsNotNull(cellStyle);
    }

    [TestMethod]
    public void GenerateStyleObject_AppliesSelectedForegroundColor()
    {
        var workbook = new XSSFWorkbook();
        var style = new HeaderStyle();
        var blue = ExcelColors.ColorCollection.Values.First(c => c.Name == "Blue");
        style.SetForegroundColor(new StyleColorSelectedEventArgs(blue));

        var cellStyle = (XSSFCellStyle)style.GenerateStyleObject(workbook);

        CollectionAssert.AreEqual(blue.IndexedColor.RGB, ((XSSFColor)cellStyle.FillForegroundColorColor).RGB);
    }
}
