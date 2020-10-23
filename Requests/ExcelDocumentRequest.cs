using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExportToExcel
{
    /// <summary>
    /// Request object used to encapsulate parameters to generate response
    /// </summary>
    /// <typeparam name="T">Type of data to be exported</typeparam>
    public class ExcelDocumentRequest<T>
    {
        /// <summary>
        /// Creates a request object thats used to export data into a spreadsheet
        /// </summary>
        /// <param name="FileName">Filename of the spreadsheet</param>
        /// <param name="ItemsToExport">Data to be exported</param>
        /// <param name="HeaderStyle">Header row's style</param>
        /// <param name="BodyStyle">Style of the body rows</param>
        public ExcelDocumentRequest(string FileName, IEnumerable<T> ItemsToExport, HeaderStyle HeaderStyle = null, BodyStyle BodyStyle = null)
        {
            Workbook = new XSSFWorkbook();

            this.ItemsToExport = ItemsToExport ?? throw new ArgumentNullException($"{nameof(ItemsToExport)}: you provided nothing to export");
            this.HeaderStyle = HeaderStyle ?? new HeaderStyle();
            this.BodyStyle = BodyStyle ?? new BodyStyle();


            if (MimeMapping.IsExtensionMissing(FileName))
                this.FileName = FileName += ".xlsx";
            else
                this.FileName = FileName;

            this.HeaderStyle.GenerateStyleObject(Workbook);
            this.BodyStyle.GenerateStyleObject(Workbook);
        }

        public Type ObjectType { get; set; }

        /// <summary>
        /// Filename of the spreadsheet
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Data that will be in the spreadsheet
        /// </summary>
        public IEnumerable<T> ItemsToExport { get; set; } = Enumerable.Empty<T>();

        /// <summary>
        /// Style for the header of the spreadsheet
        /// </summary>
        public HeaderStyle HeaderStyle { get; } = new HeaderStyle();

        /// <summary>
        /// Style for the body of the spreadsheet
        /// </summary>
        public BodyStyle BodyStyle { get; } = new BodyStyle();

        /// <summary>
        /// Instance of the Workbook
        /// </summary>
        public IWorkbook Workbook { get; }
    }
}
