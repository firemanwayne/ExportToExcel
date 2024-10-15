namespace Simple.ExportToExcel.Delegates;
using System.Threading.Tasks;

public delegate Task<UploadResponse> UploadFileEventHandler(ExcelDocumentResponse response);
