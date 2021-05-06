namespace Simple.ExportToExcel
{
    public abstract record UploadResponse
    {
        public string FileName { get; init; }
        public string ContentType { get; init; }
    }
}