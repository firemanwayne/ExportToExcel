﻿using System;

namespace ExportToExcel
{
    /// <summary>
    /// Response returned after a spreadsheet has been generated
    /// </summary>
    public class ExcelDocumentResponse
    {
        /// <summary>
        /// Creates a response object after a spreadsheet has been generated
        /// </summary>
        /// <param name="FileName">Filename of the spreadsheet</param>
        public ExcelDocumentResponse(string FileName)
        {
            this.FileName = FileName;
            ContentType = ExcelConstants.ContentType;
            IsSuccessful = true;
        }

        public ExcelDocumentResponse(Exception ex)
        {
            Exception = ex;
            IsSuccessful = false;
        }

        /// <summary>
        /// Filename for the spreadsheet
        /// </summary>
        public string FileName { get; }

        /// <summary>
        /// Content type of the spreadsheet
        /// </summary>
        public string ContentType { get; }

        /// <summary>
        /// Spreadsheet in bytes
        /// </summary>
        public byte[] SpreadSheetBytes { get; set; }

        /// <summary>
        /// Whether the operation was successful
        /// </summary>
        public bool IsSuccessful { get; }

        public Exception Exception { get; }
    }
}