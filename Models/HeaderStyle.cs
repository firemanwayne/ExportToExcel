using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;

namespace ExportToExcel
{
    public class HeaderStyle
    {
        private IWorkbook workbook;

        /// <summary>
        /// Foreground Color
        /// </summary>
        public HSSFColor ForegroundColor { get; set; } = new HSSFColor.Black();

        /// <summary>
        /// Background Color. Default is white
        /// </summary>
        public HSSFColor BackgroundColor { get; set; } = new HSSFColor.White();

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
        public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.CenterSelection;

        public ICellStyle HeaderCellStyle { get; private set; }

        public void ChangeBackgroundColor(HSSFColor Color)
        {
            ForegroundColor = Color;
            GenerateStyleObject(workbook);
        }

        public ICellStyle GenerateStyleObject(IWorkbook workbook)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));

            HeaderCellStyle = workbook.CreateCellStyle();
            HeaderCellStyle.BorderTop = TopStyle;
            HeaderCellStyle.BorderRight = RightStyle;
            HeaderCellStyle.BorderBottom = BottomStyle;
            HeaderCellStyle.BorderLeft = LeftStyle;
            HeaderCellStyle.FillForegroundColor = ForegroundColor.Indexed;
            HeaderCellStyle.FillPattern = FillPattern;
            HeaderCellStyle.FillBackgroundColor = BackgroundColor.Indexed;

            return HeaderCellStyle;
        }
    }
}
