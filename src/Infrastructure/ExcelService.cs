using System;
using System.IO;
using System.Threading.Tasks;

namespace Simple.ExportToExcel
{
    internal class ExcelService<T> : IExportToExcel<T>
    {
        public ExcelService() { }

        public Task<ExcelDocumentResponse> ExportToExcel(ExcelDocumentRequest<T> Request)
        {
            try
            {
                BuildSpreadsheet(Request);

                using var ms = new MemoryStream();
                Request.Workbook.Write(ms);

                return Task.FromResult(
                    new ExcelDocumentResponse(Request.FileName)
                    {
                        SpreadSheetBytes = ms.ToArray(),
                    });
            }
            catch (Exception ex)
            {
                return Task.FromResult(new ExcelDocumentResponse(ex));
            }
        }

        static void BuildSpreadsheet(ExcelDocumentRequest<T> Request)
        {
            var header = new HeaderBuilder<T>(Request);

            var body = new BodyBuilder<T>(Request, header);

            ExcelBuilder.Build(Request.FileName, body, header);
        }
    }
}