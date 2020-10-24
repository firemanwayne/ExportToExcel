using System.Threading.Tasks;

namespace ExportToExcel.Delegates
{
    public delegate Task<UploadResponse> UploadFileEventHandler(ExcelDocumentResponse response);
}
