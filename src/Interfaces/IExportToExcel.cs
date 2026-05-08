namespace Simple.ExportToExcel;

/// <summary>
/// Generates Excel spreadsheets
/// </summary>
/// <typeparam name="T">The type of entity to report will generate</typeparam>
public interface IExportToExcel<T>
{
    /// <summary>
    /// Generates an Excel spreadsheet from the data in <paramref name="request"/> and
    /// returns the result as a byte array wrapped in an <see cref="ExcelDocumentResponse"/>.
    /// </summary>
    /// <param name="request">The request containing the data, file name, and style configuration.</param>
    /// <returns>
    /// An <see cref="ExcelDocumentResponse"/> containing the spreadsheet bytes on success,
    /// or the caught exception on failure.
    /// </returns>
    Task<ExcelDocumentResponse> ExportToExcel(ExcelDocumentRequest<T> request);
}