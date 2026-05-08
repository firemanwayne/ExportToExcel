using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class RowStyleTests
{
    [TestMethod]
    public void DefaultForegroundColor_IsWhite()
    {
        var style = new RowStyle();

        Assert.AreEqual(HSSFColor.White.Index, style.ForegroundColor);
    }

    [TestMethod]
    public void DefaultForegroundPattern_IsSolidForeground()
    {
        var style = new RowStyle();

        Assert.AreEqual(FillPattern.SolidForeground, style.ForegroundPattern);
    }

    [TestMethod]
    public void DefaultBorders_AreThin()
    {
        var style = new RowStyle();

        Assert.AreEqual(BorderStyle.Thin, style.TopStyle);
        Assert.AreEqual(BorderStyle.Thin, style.RightStyle);
        Assert.AreEqual(BorderStyle.Thin, style.BottomStyle);
        Assert.AreEqual(BorderStyle.Thin, style.LeftStyle);
    }

    [TestMethod]
    public void RowCellStyle_IsNullBeforeGenerate()
    {
        var style = new RowStyle();

        Assert.IsNull(style.RowCellStyle);
    }

    [TestMethod]
    public void GenerateStyleObject_ReturnsNonNull()
    {
        var workbook = new XSSFWorkbook();
        var style = new RowStyle();

        var cellStyle = style.GenerateStyleObject(workbook);

        Assert.IsNotNull(cellStyle);
    }

    [TestMethod]
    public void GenerateStyleObject_SetsRowCellStyle()
    {
        var workbook = new XSSFWorkbook();
        var style = new RowStyle();

        style.GenerateStyleObject(workbook);

        Assert.IsNotNull(style.RowCellStyle);
    }

    [TestMethod]
    public void GenerateStyleObject_AppliesBorderStyles()
    {
        var workbook = new XSSFWorkbook();
        var style = new RowStyle { TopStyle = BorderStyle.Thick, BottomStyle = BorderStyle.Dashed };

        var cellStyle = style.GenerateStyleObject(workbook);

        Assert.AreEqual(BorderStyle.Thick,  cellStyle.BorderTop);
        Assert.AreEqual(BorderStyle.Dashed, cellStyle.BorderBottom);
    }

    [TestMethod]
    public void GenerateStyleObject_AppliesFillPattern()
    {
        var workbook = new XSSFWorkbook();
        var style = new RowStyle { ForegroundPattern = FillPattern.AltBars };

        var cellStyle = style.GenerateStyleObject(workbook);

        Assert.AreEqual(FillPattern.AltBars, cellStyle.FillPattern);
    }
}
