using System;
using System.Threading.Tasks;

namespace Simple.ExportToExcel
{
    /// <summary>
    /// Download exported excel File to the browser
    /// </summary>
    internal interface IExcelDownloadService : IAsyncDisposable
    {
        /// <summary>
        /// Downloads the excel file to the browser
        /// </summary>
        /// <param name="Response"></param>
        /// <returns></returns>
        ValueTask ExportFile(UploadResponse Response);
    }
}
