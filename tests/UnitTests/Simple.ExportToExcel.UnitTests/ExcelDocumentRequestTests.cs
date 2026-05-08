namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class ExcelDocumentRequestTests
{
    [TestMethod]
    public void FileName_WithoutExtension_AppendsXlsx()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("MyReport", []);

        Assert.AreEqual("MyReport.xlsx", request.FileName);
    }

    [TestMethod]
    public void FileName_WithXlsxExtension_IsPreserved()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("MyReport.xlsx", []);

        Assert.AreEqual("MyReport.xlsx", request.FileName);
    }

    [TestMethod]
    public void NullItems_ThrowsArgumentNullException()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() =>
            new ExcelDocumentRequest<DisplayModel>("Report", null));
    }

    [TestMethod]
    public void Workbook_IsInitialized()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("Report", []);

        Assert.IsNotNull(request.Workbook);
    }

    [TestMethod]
    public void HeaderStyle_DefaultsToNonNull()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("Report", []);

        Assert.IsNotNull(request.HeaderStyle);
    }

    [TestMethod]
    public void BodyStyle_DefaultsToNonNull()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("Report", []);

        Assert.IsNotNull(request.BodyStyle);
    }

    [TestMethod]
    public void CustomHeaderStyle_IsUsed()
    {
        var headerStyle = new HeaderStyle();
        var request = new ExcelDocumentRequest<DisplayModel>("Report", [], headerStyle);

        Assert.AreSame(headerStyle, request.HeaderStyle);
    }

    [TestMethod]
    public void CustomBodyStyle_IsUsed()
    {
        var bodyStyle = new BodyStyle();

        var request = new ExcelDocumentRequest<DisplayModel>("Report", [], null, bodyStyle);

        Assert.AreSame(bodyStyle, request.BodyStyle);
    }
}
