using System.IO;

namespace Simple.ExportToExcel;

/// <summary>
/// Internal implementation of <see cref="IExportToExcel{T}"/> that orchestrates building
/// and serializing an Excel spreadsheet to a byte array.
/// </summary>
/// <typeparam name="T">The entity type to export.</typeparam>
internal class ExcelService<T> : IExportToExcel<T>
{
    /// <summary>
    /// Initializes a new instance of <see cref="ExcelService{T}"/>.
    /// </summary>
    public ExcelService() { }

    /// <inheritdoc/>
    public Task<ExcelDocumentResponse> ExportToExcel(ExcelDocumentRequest<T> Request)
    {
        try
        {
            BuildSpreadsheet(Request);

            using MemoryStream ms = new();
            Request.Workbook.Write(ms);

            return Task.FromResult(
                new ExcelDocumentResponse(Request.FileName)
                {
                    SpreadSheetBytes = ms.ToArray(),
                });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ExcelDocumentResponse(ex));
        }
    }

    /// <summary>
    /// Creates the <see cref="HeaderBuilder{T}"/> and <see cref="BodyBuilder{T}"/> instances
    /// and delegates to <see cref="ExcelBuilder"/> to assemble the worksheet.
    /// </summary>
    static void BuildSpreadsheet(ExcelDocumentRequest<T> Request)
    {
        HeaderBuilder<T> header = new(Request);

        BodyBuilder<T> body = new(Request, header);

        ExcelBuilder.Build(Request.FileName, body, header);
    }
}