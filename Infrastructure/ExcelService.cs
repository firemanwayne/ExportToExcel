using System;
using System.IO;
using System.Threading.Tasks;

namespace ExportToExcel
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

        private static void BuildSpreadsheet(ExcelDocumentRequest<T> Request)
        {
            var header = new HeaderBuilder<T>(Request.HeaderStyle, Request.Workbook);
            var body = new BodyBuilder<T>(Request.ItemsToExport, Request.BodyStyle, Request.Workbook, Request.HeaderStyle);

            ExcelBuilder.Build<T>(Request.Workbook, Request.FileName, body, header);
        }
    }
}