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
    public async Task<ExcelDocumentResponse> ExportToExcel(ExcelDocumentRequest<T> Request)
    {
        try
        {
            BuildSpreadsheet(Request);

            using MemoryStream ms = new();
            await Task.Run(() => Request.Workbook.Write(ms));

            return new ExcelDocumentResponse(Request.FileName)
            {
                SpreadSheetBytes = ms.ToArray(),
            };
        }
        catch (Exception ex)
        {
            return new ExcelDocumentResponse(ex);
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