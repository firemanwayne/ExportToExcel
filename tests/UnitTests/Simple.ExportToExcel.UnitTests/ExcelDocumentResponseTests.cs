namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class ExcelDocumentResponseTests
{
    [TestMethod]
    public void SuccessConstructor_IsSuccessful_IsTrue()
    {
        var response = new ExcelDocumentResponse("Report.xlsx");

        Assert.IsTrue(response.IsSuccessful);
    }

    [TestMethod]
    public void SuccessConstructor_FileName_IsSet()
    {
        var response = new ExcelDocumentResponse("Report.xlsx");

        Assert.AreEqual("Report.xlsx", response.FileName);
    }

    [TestMethod]
    public void SuccessConstructor_ContentType_IsExcelMime()
    {
        var response = new ExcelDocumentResponse("Report.xlsx");

        Assert.AreEqual(ExcelConstants.Excel, response.ContentType);
    }

    [TestMethod]
    public void SuccessConstructor_Exception_IsNull()
    {
        var response = new ExcelDocumentResponse("Report.xlsx");

        Assert.IsNull(response.Exception);
    }

    [TestMethod]
    public void FailureConstructor_IsSuccessful_IsFalse()
    {
        var response = new ExcelDocumentResponse(new InvalidOperationException("oops"));

        Assert.IsFalse(response.IsSuccessful);
    }

    [TestMethod]
    public void FailureConstructor_Exception_IsSet()
    {
        var ex = new InvalidOperationException("oops");
        var response = new ExcelDocumentResponse(ex);

        Assert.AreSame(ex, response.Exception);
    }

    [TestMethod]
    public void FailureConstructor_FileName_IsNull()
    {
        var response = new ExcelDocumentResponse(new Exception());

        Assert.IsNull(response.FileName);
    }

    [TestMethod]
    public void FailureConstructor_ContentType_IsNull()
    {
        var response = new ExcelDocumentResponse(new Exception());

        Assert.IsNull(response.ContentType);
    }

    [TestMethod]
    public void SpreadSheetBytes_DefaultsToNull()
    {
        var response = new ExcelDocumentResponse("Report.xlsx");

        Assert.IsNull(response.SpreadSheetBytes);
    }

    [TestMethod]
    public void SpreadSheetBytes_CanBeSet()
    {
        var response = new ExcelDocumentResponse("Report.xlsx");
        response.SpreadSheetBytes = new byte[] { 1, 2, 3 };

        CollectionAssert.AreEqual(new byte[] { 1, 2, 3 }, response.SpreadSheetBytes);
    }
}
