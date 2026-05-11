namespace Simple.ExportToExcel;

/// <summary>
/// Response returned after a spreadsheet has been generated
/// </summary>
public class ExcelDocumentResponse
{
    /// <summary>
    /// Creates a response object after a spreadsheet has been generated
    /// </summary>
    /// <param name="fileName">Filename of the spreadsheet</param>
    public ExcelDocumentResponse(string fileName)
    {
        FileName = fileName;
        ContentType = ExcelConstants.Excel;
        IsSuccessful = true;
    }

    /// <summary>
    /// Creates a failure response capturing the exception that caused the export to fail.
    /// </summary>
    /// <param name="ex">The exception that occurred during export.</param>
    public ExcelDocumentResponse(Exception ex)
    {
        Exception = ex;
        IsSuccessful = false;
    }

    /// <summary>
    /// Filename for the spreadsheet
    /// </summary>
    public string FileName { get; }

    /// <summary>
    /// Content type of the spreadsheet
    /// </summary>
    public string ContentType { get; }

    /// <summary>
    /// Spreadsheet in bytes
    /// </summary>
    public byte[] SpreadSheetBytes { get; set; }

    /// <summary>
    /// Whether the operation was successful
    /// </summary>
    public bool IsSuccessful { get; }

    /// <summary>
    /// The exception that caused the export to fail, or <c>null</c> if the operation succeeded.
    /// </summary>
    public Exception Exception { get; }
}