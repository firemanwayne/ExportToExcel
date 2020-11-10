namespace ExportToExcel
{
    public record UploadFileLocalResponse : UploadResponse
    {
        public string FileContent { get; init; }
    }
}
