﻿namespace Simple.ExportToExcel;

public class ExcelBuilder
{
    private ExcelBuilder()
    { }

    /// <summary>
    /// Builds the Excel Spreadsheet from the Header Builder and BodyBuilder classes
    /// </summary>
    /// <typeparam name="T">Type of Data Object</typeparam>
    /// <param name="WorkBook">Instance of the Workbook</param>
    /// <param name="FileName">Spreadsheet's file name</param>
    /// <param name="Body">Body of the spreadsheet</param>
    /// <param name="Header">Header of the Spreadsheet</param>
    public static void Build<T>(string FileName, BodyBuilder<T> Body, HeaderBuilder<T> Header)
    {
        NPOI.SS.UserModel.ISheet Sheet = Header.Build(FileName);
        Body.Build(Sheet);
    }
}
