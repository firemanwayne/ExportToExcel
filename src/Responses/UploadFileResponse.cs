using System;

namespace Simple.ExportToExcel;

public record UploadFileResponse : UploadResponse
{
    public Uri FileUri { get; init; }
}
