namespace Simple.ExportToExcel;

/// <summary>
/// Download exported excel File to the browser
/// </summary>
internal interface IExcelDownloadService : IAsyncDisposable
{
    /// <summary>
    /// Triggers a file download in the browser based on the type of <see cref="UploadResponse"/>.
    /// Uses a URI-based download for <see cref="UploadFileResponse"/> or inline content for
    /// <see cref="UploadFileLocalResponse"/>.
    /// </summary>
    /// <param name="Response">The upload response containing the file data or URI.</param>
    /// <returns>A <see cref="ValueTask"/> representing the asynchronous operation.</returns>
    ValueTask ExportFile(UploadResponse Response);
}
