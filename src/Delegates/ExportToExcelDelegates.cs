namespace Simple.ExportToExcel.Delegates;

/// <summary>
/// Delegate invoked after an Excel file has been generated. Implementations are
/// responsible for uploading or otherwise persisting the file and returning an
/// <see cref="UploadResponse"/> describing the result.
/// </summary>
/// <param name="response">The generated Excel document response containing the file bytes.</param>
/// <returns>An <see cref="UploadResponse"/> describing the upload result.</returns>
public delegate Task<UploadResponse> UploadFileEventHandler(ExcelDocumentResponse response);
