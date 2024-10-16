﻿using System;
using System.IO;
using System.Threading.Tasks;

namespace Simple.ExportToExcel;

internal class ExcelService<T> : IExportToExcel<T>
{
    public ExcelService() { }

    public Task<ExcelDocumentResponse> ExportToExcel(ExcelDocumentRequest<T> Request)
    {
        try
        {
            BuildSpreadsheet(Request);

            using MemoryStream ms = new();
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
        HeaderBuilder<T> header = new(Request);

        BodyBuilder<T> body = new(Request, header);

        ExcelBuilder.Build(Request.FileName, body, header);
    }
}