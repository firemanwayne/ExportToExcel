using Microsoft.JSInterop;

namespace Simple.ExportToExcel;

/// <summary>
/// Internal implementation of <see cref="IExcelDownloadService"/> that uses JavaScript
/// interop to trigger a file download in the browser.
/// </summary>
internal class ExcelDownloadService : IExcelDownloadService
{
    static string ExportUriFunctionName => "ExportFileToUri";

    static string ExportLocalFunctionName => "ExportFile";

    readonly Lazy<Task<IJSObjectReference>> _moduleTask;

    /// <summary>
    /// Initializes a new <see cref="ExcelDownloadService"/> with the given JS runtime.
    /// The JavaScript module is loaded lazily on first use.
    /// </summary>
    /// <param name="JS">The Blazor JS interop runtime.</param>
    public ExcelDownloadService(IJSRuntime JS)
        => _moduleTask = new(() => JS.InvokeAsync<IJSObjectReference>(
            "import", "./_content/Simple.ExportToExcel/ToExcel.js").AsTask());

    public async ValueTask ExportFile(UploadResponse Response)
    {
        IJSObjectReference module = await _moduleTask.Value;

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
        if (_moduleTask.IsValueCreated)
        {
            IJSObjectReference module = await _moduleTask.Value;
            await module.DisposeAsync();
        }
    }
}
