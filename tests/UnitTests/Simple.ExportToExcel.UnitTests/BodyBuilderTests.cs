using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class BodyBuilderTests
{
    static ISheet BuildSheet<T>(ExcelDocumentRequest<T> request)
    {
        var header = new HeaderBuilder<T>(request);
        var body = new BodyBuilder<T>(request, header);
        ExcelBuilder.Build(request.FileName, body, header);
        return request.Workbook.GetSheet(request.FileName);
    }

    [TestMethod]
    public void Build_StringCell_HasCorrectValue()
    {
        var data = new[] { new DisplayModel { FirstName = "Alice", Age = 30, Score = 95.5 } };
        var request = new ExcelDocumentRequest<DisplayModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("Alice", sheet.GetRow(1).GetCell(0).StringCellValue);
    }

    [TestMethod]
    public void Build_IntCell_HasCorrectValue()
    {
        var data = new[] { new DisplayModel { FirstName = "Alice", Age = 42, Score = 0 } };
        var request = new ExcelDocumentRequest<DisplayModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual(42, (int)sheet.GetRow(1).GetCell(1).NumericCellValue);
    }

    [TestMethod]
    public void Build_DoubleCell_HasCorrectValue()
    {
        var data = new[] { new DisplayModel { FirstName = "Alice", Age = 0, Score = 88.5 } };
        var request = new ExcelDocumentRequest<DisplayModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual(88.5, sheet.GetRow(1).GetCell(2).NumericCellValue);
    }

    [TestMethod]
    public void Build_FloatCell_HasCorrectValue()
    {
        var data = new[] { new FloatModel { Ratio = 3.14f, Label = "pi" } };
        var request = new ExcelDocumentRequest<FloatModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual((double)3.14f, sheet.GetRow(1).GetCell(0).NumericCellValue, 0.0001);
    }

    [TestMethod]
    public void Build_FallbackCell_WritesStringRepresentation()
    {
        var data = new[] { new MixedTypeModel { Text = "hello", Flag = true } };
        var request = new ExcelDocumentRequest<MixedTypeModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("True", sheet.GetRow(1).GetCell(1).StringCellValue);
    }

    [TestMethod]
    public void Build_NullPropertyValue_WritesEmptyString()
    {
        var data = new[] { new PlainModel { Name = null!, Value = 0 } };
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("", sheet.GetRow(1).GetCell(0).StringCellValue);
    }

    [TestMethod]
    public void Build_MultipleRows_AllRowsPresent()
    {
        var data = new[]
        {
            new PlainModel { Name = "A", Value = 1 },
            new PlainModel { Name = "B", Value = 2 },
            new PlainModel { Name = "C", Value = 3 },
        };
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.IsNotNull(sheet.GetRow(1));
        Assert.IsNotNull(sheet.GetRow(2));
        Assert.IsNotNull(sheet.GetRow(3));
        Assert.IsNull(sheet.GetRow(4));
    }

    [TestMethod]
    public void Build_EmptyData_OnlyHasHeaderRow()
    {
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", []);

        var sheet = BuildSheet(request);

        Assert.IsNotNull(sheet.GetRow(0));
        Assert.IsNull(sheet.GetRow(1));
    }

    [TestMethod]
    public void Build_CellsHaveBodyStyle()
    {
        var data = new[] { new PlainModel { Name = "Test", Value = 1 } };
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.IsNotNull(sheet.GetRow(1).GetCell(0).CellStyle);
    }

    [TestMethod]
    public void DataItems_ReflectsRequestItems()
    {
        var data = new[] { new PlainModel { Name = "Test", Value = 1 } };
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", data);
        var header = new HeaderBuilder<PlainModel>(request);
        var body = new BodyBuilder<PlainModel>(request, header);

        Assert.AreSame(data, body.DataItems);
    }

    [TestMethod]
    public void BodyStyle_ReflectsRequestBodyStyle()
    {
        var bodyStyle = new BodyStyle();
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", [], null, bodyStyle);
        var header = new HeaderBuilder<PlainModel>(request);
        var body = new BodyBuilder<PlainModel>(request, header);

        Assert.AreSame(bodyStyle, body.BodyStyle);
    }

    [TestMethod]
    public void Build_MultipleRows_CorrectValuesInEachRow()
    {
        var data = new[]
        {
            new PlainModel { Name = "First", Value = 10 },
            new PlainModel { Name = "Second", Value = 20 },
        };
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("First",  sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.AreEqual("Second", sheet.GetRow(2).GetCell(0).StringCellValue);
        Assert.AreEqual(10, (int)sheet.GetRow(1).GetCell(1).NumericCellValue);
        Assert.AreEqual(20, (int)sheet.GetRow(2).GetCell(1).NumericCellValue);
    }

    // --- Attribute-based column layout ---

    [TestMethod]
    public void Build_WithAttributeModel_ColumnAWrittenAtIndex0()
    {
        var data = new[] { new AttributeModel { A = "alpha", B = "beta" } };
        var request = new ExcelDocumentRequest<AttributeModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("alpha", sheet.GetRow(1).GetCell(0).StringCellValue);
    }

    [TestMethod]
    public void Build_WithAttributeModel_ColumnBWrittenAtIndex1()
    {
        var data = new[] { new AttributeModel { A = "alpha", B = "beta" } };
        var request = new ExcelDocumentRequest<AttributeModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("beta", sheet.GetRow(1).GetCell(1).StringCellValue);
    }

    [TestMethod]
    public void Build_WithAttributeModel_MultipleRowsHaveCorrectValues()
    {
        var data = new[]
        {
            new AttributeModel { A = "a1", B = "b1" },
            new AttributeModel { A = "a2", B = "b2" },
        };
        var request = new ExcelDocumentRequest<AttributeModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("a1", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.AreEqual("b1", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.AreEqual("a2", sheet.GetRow(2).GetCell(0).StringCellValue);
        Assert.AreEqual("b2", sheet.GetRow(2).GetCell(1).StringCellValue);
    }

    // --- Generic property filtering ---

    [TestMethod]
    public void Build_GenericProperties_AreNotWrittenAsCells()
    {
        // ListPropertyModel has Name (string), Items (IList<string>), Numbers (IEnumerable<int>).
        // CreateRow filters out generic-type properties, so only Name gets a cell.
        var data = new[] { new ListPropertyModel { Name = "only-scalar" } };
        var request = new ExcelDocumentRequest<ListPropertyModel>("Sheet", data);

        var sheet = BuildSheet(request);

        Assert.AreEqual("only-scalar", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.IsNull(sheet.GetRow(1).GetCell(1)); // Items — skipped
        Assert.IsNull(sheet.GetRow(1).GetCell(2)); // Numbers — skipped
    }

    [TestMethod]
    public void Build_OnlyScalarColumnCount_MatchesNonGenericPropertyCount()
    {
        // Verify that BodyBuilder only writes cells for non-generic properties.
        var data = new[] { new ListPropertyModel { Name = "test" } };
        var request = new ExcelDocumentRequest<ListPropertyModel>("Sheet", data);

        var sheet = BuildSheet(request);

        int cellCount = sheet.GetRow(1).LastCellNum; // 1-based count of last cell
        Assert.AreEqual(1, cellCount);
    }

    // --- Row indexing ---

    [TestMethod]
    public void Build_FirstDataRow_IsAtSheetRowIndex1()
    {
        var data = new[] { new PlainModel { Name = "row-one", Value = 1 } };
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", data);

        var sheet = BuildSheet(request);

        // Row 0 is the header — its first cell holds the column name, not the data value.
        Assert.AreEqual("Name",    sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.AreEqual("row-one", sheet.GetRow(1).GetCell(0).StringCellValue);
    }

    // --- Style application ---

    [TestMethod]
    public void Build_CustomBodyStyle_IsAppliedToDataCells()
    {
        var bodyStyle = new BodyStyle();
        var data = new[] { new PlainModel { Name = "styled", Value = 99 } };
        var request = new ExcelDocumentRequest<PlainModel>("Sheet", data, null, bodyStyle);

        var sheet = BuildSheet(request);

        // Both cells in the row should carry the body cell style index.
        short styleIndex = sheet.GetRow(1).GetCell(0).CellStyle.Index;
        Assert.AreEqual(styleIndex, sheet.GetRow(1).GetCell(1).CellStyle.Index);
    }

    // --- CreateBodyRowFromGenericList dead-code coverage note ---

    [TestMethod]
    public void Build_GenericListProperty_ProducesNoAdditionalRows()
    {
        // HeaderBuilder skips IList<T> properties (does not add them to ColumnProperties),
        // so BodyBuilder never writes cells or sub-rows for them.
        var data = new[] { new ListPropertyModel
        {
            Name = "parent",
            Items = ["child-a", "child-b"],
        }};
        var request = new ExcelDocumentRequest<ListPropertyModel>("Sheet", data);

        var sheet = BuildSheet(request);

        // Only the scalar Name column is written; no sub-rows for Items.
        Assert.AreEqual(1, sheet.LastRowNum); // row 0 = header, row 1 = data, nothing more
    }
}
