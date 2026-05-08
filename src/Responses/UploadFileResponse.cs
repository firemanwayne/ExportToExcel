namespace Simple.ExportToExcel;

/// <summary>
/// Upload response that carries a remote URI pointing to the stored file.
/// </summary>
public record UploadFileResponse : UploadResponse
{
    /// <summary>
    /// URI of the remotely stored file (e.g. an Azure Blob Storage URL).
    /// </summary>
    public Uri FileUri { get; init; }
}
