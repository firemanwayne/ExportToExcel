namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class ExcelServiceTests
{
    static readonly IEnumerable<DisplayModel> SampleData =
    [
        new() { FirstName = "Alice", Age = 30, Score = 95.5 },
        new() { FirstName = "Bob",   Age = 25, Score = 88.0 },
    ];

    ExcelService<DisplayModel> _service = null!;

    [TestInitialize]
    public void Setup() => _service = new ExcelService<DisplayModel>();

    [TestMethod]
    public async Task ExportToExcel_ReturnsSuccessfulResponse()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("Report", SampleData);

        var response = await _service.ExportToExcel(request);

        Assert.IsTrue(response.IsSuccessful);
    }

    [TestMethod]
    public async Task ExportToExcel_ResponseContainsBytes()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("Report", SampleData);

        var response = await _service.ExportToExcel(request);

        Assert.IsNotNull(response.SpreadSheetBytes);
        Assert.IsTrue(response.SpreadSheetBytes.Length > 0);
    }

    [TestMethod]
    public async Task ExportToExcel_ResponseHasCorrectFileName()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("MySheet.xlsx", SampleData);

        var response = await _service.ExportToExcel(request);

        Assert.AreEqual("MySheet.xlsx", response.FileName);
    }

    [TestMethod]
    public async Task ExportToExcel_ResponseHasExcelContentType()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("Report", SampleData);

        var response = await _service.ExportToExcel(request);

        Assert.AreEqual(ExcelConstants.Excel, response.ContentType);
    }

    [TestMethod]
    public async Task ExportToExcel_WithEmptyData_ReturnsSuccessfulResponse()
    {
        var request = new ExcelDocumentRequest<DisplayModel>("Report", []);

        var response = await _service.ExportToExcel(request);

        Assert.IsTrue(response.IsSuccessful);
    }

    [TestMethod]
    public async Task ExportToExcel_WithAttributeModel_ReturnsSuccessfulResponse()
    {
        var service = new ExcelService<AttributeModel>();
        var data = new[] { new AttributeModel { A = "a1", B = "b1" } };
        var request = new ExcelDocumentRequest<AttributeModel>("Report", data);

        var response = await service.ExportToExcel(request);

        Assert.IsTrue(response.IsSuccessful);
    }

    [TestMethod]
    public async Task ExportToExcel_WithPlainModel_ReturnsSuccessfulResponse()
    {
        var service = new ExcelService<PlainModel>();
        var data = new[] { new PlainModel { Name = "Test", Value = 42 } };
        var request = new ExcelDocumentRequest<PlainModel>("Report", data);

        var response = await service.ExportToExcel(request);

        Assert.IsTrue(response.IsSuccessful);
    }
}
