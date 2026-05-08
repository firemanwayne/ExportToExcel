namespace Simple.ExportToExcel;

/// <summary>
/// Upload response that carries the file content as a base64-encoded string
/// for in-browser download without a remote URI.
/// </summary>
public record UploadFileLocalResponse : UploadResponse
{
    /// <summary>
    /// Base64-encoded content of the uploaded file.
    /// </summary>
    public string FileContent { get; init; }
}
