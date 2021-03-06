﻿using ExportToExcel.Styles;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace ExportToExcel
{
    public abstract class ExcelStyle
    {
        short foregroundColorIndex;
        short backgroundColorIndex;
        static readonly IDictionary<short, ExcelColors> colors = new Dictionary<short, ExcelColors>();

        public ExcelStyle()
        {
            foreach (var item in ExcelColors.ColorCollection)
                colors.TryAdd(item.Key, item.Value);
        }

        public IEnumerable<ExcelColors> Colors { get => colors.Values; }

        public ICellStyle CellStyle { get; private set; }

        public ICellStyle GenerateStyleObject(in IWorkbook workbook)
        {
            var cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();

            cellStyle.BorderTop = TopStyle;
            cellStyle.BorderRight = RightStyle;
            cellStyle.BorderBottom = BottomStyle;
            cellStyle.BorderLeft = LeftStyle;
            cellStyle.FillPattern = FillPattern;

            cellStyle.VerticalAlignment = VerticalAlignment;
            cellStyle.Alignment = HorizontalAlignment;

            cellStyle.SetFillForegroundColor(ForegroundColor);
            cellStyle.SetFillBackgroundColor(BackgroundColor);

            CellStyle = cellStyle;

            return cellStyle;
        }

        /// <summary>
        /// Foreground Color
        /// </summary>
        public short ForegroundColorIndex
        {
            get => foregroundColorIndex;
            set
            {
                foregroundColorIndex = value;

                if (colors.TryGetValue(value, out var r))
                    ForegroundColor = new XSSFColor(r.IndexedColor);
            }
        }

        /// <summary>
        /// Background Color. Default is white
        /// </summary>
        public short BackgroundColorIndex
        {
            get => backgroundColorIndex;
            set
            {
                backgroundColorIndex = value;

                if (colors.TryGetValue(value, out var r))
                    BackgroundColor = new XSSFColor(r.IndexedColor);
            }
        }

        /// <summary>
        /// Foreground Color
        /// </summary>
        public XSSFColor ForegroundColor { get; private set; }

        /// <summary>
        /// Background Color. Default is white
        /// </summary>
        public XSSFColor BackgroundColor { get; private set; }

        /// <summary>
        /// Vertical Alignment
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Center;

        /// <summary>
        /// Horizontal Alignment
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.Center;

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

        public void SetForegroundColor(StyleColorSelectedEventArgs a)
        {
            ForegroundColorIndex = a.ColorIndex;
        }

        public void SetHorizontalAlignment(HorizontalAlignmentChangedEventArgs a)
        {
            HorizontalAlignment = a.SelectedAlignment;
        }

        public void SetVerticalAlignment(VerticalAlignmentChangedEventArgs a)
        {
            VerticalAlignment = a.SelectedAlignment;
        }
    }
}
