using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Drawing;

namespace ExportToExcel
{
    public class BodyStyle
    {
        private IWorkbook workbook;

        /// <summary>
        /// Foreground Color. Text Color
        /// </summary>
        public XSSFColor ForegroundColor { get; set; } = new XSSFColor(Color.Black);

        /// <summary>
        /// Background Color. Default is white
        /// </summary>
        public XSSFColor BackgroundColor { get; set; } = new XSSFColor(Color.Transparent);

        /// <summary>
        /// Fill Pattern
        /// </summary>
        public FillPattern FillPattern { get; set; } = FillPattern.SolidForeground;

        /// <summary>
        /// Top Border
        /// </summary>
        public BorderStyle TopStyle { get; set; } = BorderStyle.Thin;

        /// <summary>
        /// Right Border
        /// </summary>
        public BorderStyle RightStyle { get; set; } = BorderStyle.Thin;

        /// <summary>
        /// Bottom Border
        /// </summary>
        public BorderStyle BottomStyle { get; set; } = BorderStyle.Thin;

        /// <summary>
        /// Left Border
        /// </summary>
        public BorderStyle LeftStyle { get; set; } = BorderStyle.Thin;

        /// <summary>
        /// Vertical Alignment
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Center;

        /// <summary>
        /// Horizontal Alignment
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.Center;

        public ICellStyle BodyCellStyle { get; private set; }        

        public ICellStyle GenerateStyleObject(in IWorkbook workbook)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));

            BodyCellStyle = workbook.CreateCellStyle();
            BodyCellStyle.BorderTop = TopStyle;
            BodyCellStyle.BorderRight = RightStyle;
            BodyCellStyle.BorderBottom = BottomStyle;
            BodyCellStyle.BorderLeft = LeftStyle;
            BodyCellStyle.FillForegroundColor = ForegroundColor.Indexed;
            BodyCellStyle.FillPattern = FillPattern;
            BodyCellStyle.FillBackgroundColor = BackgroundColor.Indexed;

            return BodyCellStyle;
        }
    }
}
