using System.Threading.Tasks;

namespace Simple.ExportToExcel.Delegates
{
    public delegate Task<UploadResponse> UploadFileEventHandler(ExcelDocumentResponse response);
}
