using System;

namespace ExportToExcel
{
    public abstract record UploadResponse
    {
        public string FileName { get; init; }
        public string ContentType { get; init; }
    }

    public record UploadFileResponse : UploadResponse
    {
        public Uri FileUri { get; init; }

    }

    public record UploadFileLocalResponse : UploadResponse
    {
        public string FileContent { get; init; }
    }
}
