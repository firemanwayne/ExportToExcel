using NPOI.SS.UserModel;
using Simple.ExportToExcel.Styles;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class StyleColorSelectedEventArgsTests
{
    [TestMethod]
    public void Constructor_SetsColorIndex()
    {
        var color = ExcelColors.ColorCollection.Values.First(c => c.Name == "Red");
        var args = new StyleColorSelectedEventArgs(color);

        Assert.AreEqual(color.Id, args.ColorIndex);
    }

    [TestMethod]
    public void Constructor_SetsRGBValue()
    {
        var color = ExcelColors.ColorCollection.Values.First(c => c.Name == "Red");
        var args = new StyleColorSelectedEventArgs(color);

        Assert.AreEqual(color.RGBValue, args.RGBValue);
    }

    [TestMethod]
    public void Constructor_SetsColorName()
    {
        var color = ExcelColors.ColorCollection.Values.First(c => c.Name == "Red");
        var args = new StyleColorSelectedEventArgs(color);

        Assert.AreEqual("Red", args.ColorName);
    }
}

[TestClass]
public class HorizontalAlignmentChangedEventArgsTests
{
    [TestMethod]
    public void Constructor_SetsSelectedAlignment()
    {
        var args = new HorizontalAlignmentChangedEventArgs(HorizontalAlignment.Left);

        Assert.AreEqual(HorizontalAlignment.Left, args.SelectedAlignment);
    }

    [TestMethod]
    public void Constructor_Right_SetsSelectedAlignment()
    {
        var args = new HorizontalAlignmentChangedEventArgs(HorizontalAlignment.Right);

        Assert.AreEqual(HorizontalAlignment.Right, args.SelectedAlignment);
    }
}

[TestClass]
public class VerticalAlignmentChangedEventArgsTests
{
    [TestMethod]
    public void Constructor_SetsSelectedAlignment()
    {
        var args = new VerticalAlignmentChangedEventArgs(VerticalAlignment.Top);

        Assert.AreEqual(VerticalAlignment.Top, args.SelectedAlignment);
    }

    [TestMethod]
    public void Constructor_Bottom_SetsSelectedAlignment()
    {
        var args = new VerticalAlignmentChangedEventArgs(VerticalAlignment.Bottom);

        Assert.AreEqual(VerticalAlignment.Bottom, args.SelectedAlignment);
    }
}

[TestClass]
public class ExcelColorsTests
{
    [TestMethod]
    public void ColorCollection_IsNotEmpty()
    {
        Assert.IsTrue(ExcelColors.ColorCollection.Count > 0);
    }

    [TestMethod]
    public void ColorCollection_ContainsRed()
    {
        Assert.IsTrue(ExcelColors.ColorCollection.Values.Any(c => c.Name == "Red"));
    }

    [TestMethod]
    public void ColorCollection_ContainsWhite()
    {
        Assert.IsTrue(ExcelColors.ColorCollection.Values.Any(c => c.Name == "White"));
    }

    [TestMethod]
    public void Constructor_SetsName()
    {
        var color = ExcelColors.ColorCollection.Values.First(c => c.Name == "Blue");

        Assert.AreEqual("Blue", color.Name);
    }

    [TestMethod]
    public void Constructor_SetsId()
    {
        var color = ExcelColors.ColorCollection.Values.First(c => c.Name == "Red");

        Assert.AreEqual(NPOI.SS.UserModel.IndexedColors.Red.Index, color.Id);
    }

    [TestMethod]
    public void Constructor_SetsRGBValue_WithCorrectFormat()
    {
        var color = ExcelColors.ColorCollection.Values.First(c => c.Name == "Red");

        StringAssert.StartsWith(color.RGBValue, "rgb(");
        StringAssert.EndsWith(color.RGBValue, ")");
    }

    [TestMethod]
    public void Constructor_SetsIndexedColor()
    {
        var color = ExcelColors.ColorCollection.Values.First(c => c.Name == "Green");

        Assert.IsNotNull(color.IndexedColor);
    }
}
