using ExportToExcel.Models.Concrete;
using System.Threading.Tasks;

namespace ExportToExcel.Delegates
{    
    public delegate Task<UploadFileResponse> UploadFileEventHandler(ExcelDocumentResponse response);
}
