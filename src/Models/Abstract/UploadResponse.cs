namespace Simple.ExportToExcel;

/// <summary>
/// Base record for all upload responses returned after a file operation completes.
/// </summary>
public abstract record UploadResponse
{
    /// <summary>
    /// Name of the uploaded file, including extension.
    /// </summary>
    public string FileName { get; init; }

    /// <summary>
    /// MIME content type of the uploaded file.
    /// </summary>
    public string ContentType { get; init; }
}