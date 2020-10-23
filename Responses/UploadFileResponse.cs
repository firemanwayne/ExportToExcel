using System;

namespace ExportToExcel.Models.Concrete
{
    public record UploadFileResponse
    {
        public Uri FileUri { get; }

        public string FileName { get; }
    }
}
