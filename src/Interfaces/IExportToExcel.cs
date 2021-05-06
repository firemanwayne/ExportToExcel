using System.Threading.Tasks;

namespace Simple.ExportToExcel
{
    /// <summary>
    /// Generates Excel spreadsheets
    /// </summary>
    /// <typeparam name="T">The type of entity to report will generate</typeparam>
    public interface IExportToExcel<T>
    {
        /// <summary>
        /// Generates an Excel spreadsheet for a list of Entities
        /// </summary>
        /// <param name="Entities">The data to be in the spreadsheet</param>
        /// <param name="FileName">The name you would like the spreadsheet to be saved to</param>
        /// <param name="ContainerName">The name of the Azure Blob Container you want to store the generated spreadsheet</param>        
        /// <returns></returns>
        Task<ExcelDocumentResponse> ExportToExcel(ExcelDocumentRequest<T> Request);
    }
}