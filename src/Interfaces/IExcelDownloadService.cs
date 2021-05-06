using System;
using System.Threading.Tasks;

namespace Simple.ExportToExcel
{
    internal interface IExcelDownloadService : IAsyncDisposable
    {
        ValueTask ExportFile(UploadResponse Response);
    }
}
