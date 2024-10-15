namespace Simple.ExportToExcel;

public record UploadFileLocalResponse : UploadResponse
{
    public string FileContent { get; init; }
}
