using Microsoft.JSInterop;

using System;
using System.Threading.Tasks;

namespace Simple.ExportToExcel;

internal class ExcelDownloadService : IExcelDownloadService
{
    static string ExportUriFunctionName => "ExportFileToUri";
    static string ExportLocalFunctionName => "ExportFile";

    readonly Lazy<Task<IJSObjectReference>> moduleTask;

    public ExcelDownloadService(IJSRuntime JS)
        => moduleTask = new(() => JS.InvokeAsync<IJSObjectReference>(
            "import", "./_content/Simple.ExportToExcel/ToExcel.js").AsTask());

    public async ValueTask ExportFile(UploadResponse Response)
    {
        IJSObjectReference module = await moduleTask.Value;

        if (Response is UploadFileResponse f)
        {
            await module.InvokeVoidAsync(ExportUriFunctionName, f.FileName, f.FileUri);
        }
        else if (Response is UploadFileLocalResponse l)
        {
            await module.InvokeVoidAsync(ExportLocalFunctionName, l.FileName, l.FileContent, l.ContentType);
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (moduleTask.IsValueCreated)
        {
            IJSObjectReference module = await moduleTask.Value;
            await module.DisposeAsync();
        }
    }
}
