namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class MimeMappingTests
{
    [TestMethod]
    public void IsExtensionMissing_WithExtension_ReturnsFalse()
    {
        Assert.IsFalse(MimeMapping.IsExtensionMissing("report.xlsx"));
    }

    [TestMethod]
    public void IsExtensionMissing_WithoutExtension_ReturnsTrue()
    {
        Assert.IsTrue(MimeMapping.IsExtensionMissing("report"));
    }

    [TestMethod]
    public void GetMimeMappingByExtension_KnownExtension_ReturnsCorrectMime()
    {
        var mime = MimeMapping.GetMimeMappingByExtension(".gif");

        Assert.AreEqual("image/gif", mime);
    }

    [TestMethod]
    public void GetMimeMappingByExtension_XlsExtension_ReturnsExcelMime()
    {
        var mime = MimeMapping.GetMimeMappingByExtension(".xls");

        Assert.AreEqual("application/vnd.ms-excel", mime);
    }

    [TestMethod]
    public void GetMimeMappingByExtension_UnknownExtension_ReturnsOctetStream()
    {
        var mime = MimeMapping.GetMimeMappingByExtension(".unknownxyz");

        Assert.AreEqual("application/octet-stream", mime);
    }

    [TestMethod]
    public void GetMimeFromFileName_KnownExtension_ReturnsCorrectMime()
    {
        var mime = MimeMapping.GetMimeFromFileName("document.pdf");

        Assert.AreEqual("application/pdf", mime);
    }

    [TestMethod]
    public void GetMimeFromFileName_MissingExtension_ThrowsFileNameMissingExtensionException()
    {
        Assert.ThrowsExactly<FileNameMissingExtensionException>(() =>
            MimeMapping.GetMimeFromFileName("noextension"));
    }

    [TestMethod]
    public void GetAllMimeTypes_ReturnsNonEmptyDictionary()
    {
        var result = MimeMapping.GetAllMimeTypes();

        Assert.IsTrue(result.Count > 0);
    }

    [TestMethod]
    public void GetAllMimeTypes_KeysMatchExtensions()
    {
        var result = MimeMapping.GetAllMimeTypes();

        Assert.IsTrue(result.ContainsKey(".gif"));
    }

    [TestMethod]
    public void GetMimeFromFileName_MultiSegmentName_UsesLastExtension()
    {
        // e.g. "my.report.xlsx" should resolve .xlsx, not .report
        var mime = MimeMapping.GetMimeFromFileName("my.report.xlsx");

        Assert.AreEqual("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", mime);
    }

    [TestMethod]
    public void GetMimeMappingsByType_KnownType_ReturnsOnlyMatchingExtensions()
    {
        var results = MimeMapping.GetMimeMappingsByType("image/gif");

        Assert.IsTrue(results.Count > 0);
        Assert.IsTrue(results.All(ext => MimeMapping.Extensions[ext].Value == "image/gif"));
    }

    [TestMethod]
    public void GetMimeMappingsByType_UnknownType_ReturnsEmptyList()
    {
        var results = MimeMapping.GetMimeMappingsByType("application/totally-unknown-xyz");

        Assert.AreEqual(0, results.Count);
    }

    [TestMethod]
    public void Extensions_ContainsXlsEntry()
    {
        Assert.IsTrue(MimeMapping.Extensions.ContainsKey(".xls"));
    }

    [TestMethod]
    public void Extensions_ContainsXlsxEntry()
    {
        Assert.IsTrue(MimeMapping.Extensions.ContainsKey(".xlsx"));
    }

    [TestMethod]
    public void GetMimeMappingByExtension_XlsxExtension_ReturnsCorrectMime()
    {
        var mime = MimeMapping.GetMimeMappingByExtension(".xlsx");

        Assert.AreEqual("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", mime);
    }

    [TestMethod]
    public void Extensions_ContainsFallbackEntry()
    {
        Assert.IsTrue(MimeMapping.Extensions.ContainsKey(".*"));
    }
}
