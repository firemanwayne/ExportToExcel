using System;

namespace ExportToExcel
{

    public record UploadFileResponse : UploadResponse
    {
        public Uri FileUri { get; init; }

    }
}
