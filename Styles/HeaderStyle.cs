using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace ExportToExcel
{
    public class HeaderStyle
    {        
        string foregroundColorIndex;
        string backgroundColorIndex;
        static readonly IDictionary<string, Color> colors = new Dictionary<string, Color>();
        static ICollection<Color> ColorCollection = new List<Color>();

        public HeaderStyle()
        {           
            ColorCollection = GetColors();

            foreach (var item in ColorCollection)
                colors.TryAdd($"#{item.R:X2}{item.G:X2}{item.B:X2}", item);
        }

        /// <summary>
        /// Foreground Color
        /// </summary>
        public string ForegroundColorIndex 
        {
            get => foregroundColorIndex;
            set 
            {
                foregroundColorIndex = value;

                if (colors.TryGetValue(value, out var result))
                    ForegroundColor = new XSSFColor(new[] {result.R, result.G, result.B, result.A });
            }
        }

        /// <summary>
        /// Background Color. Default is white
        /// </summary>
        public string BackgroundColorIndex 
        {
            get => backgroundColorIndex;
            set 
            {
                backgroundColorIndex = value;

                if (colors.TryGetValue(value, out var result))
                    BackgroundColor = new XSSFColor(new[] { result.R, result.G, result.B, result.A });
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

        public IEnumerable<Color> Colors { get => colors.Values; }

        public ICellStyle HeaderCellStyle { get; private set; }        

        public ICellStyle GenerateStyleObject(IWorkbook workbook)
        {         
            var cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();

            cellStyle.BorderTop = TopStyle;
            cellStyle.BorderRight = RightStyle;
            cellStyle.BorderBottom = BottomStyle;
            cellStyle.BorderLeft = LeftStyle;
            cellStyle.FillPattern = FillPattern;      

            cellStyle.SetFillForegroundColor(ForegroundColor);            
            cellStyle.SetFillBackgroundColor(BackgroundColor);

            HeaderCellStyle = cellStyle;

            return cellStyle;
        }

        static ICollection<Color> GetColors()
           => typeof(Color).GetProperties(
               BindingFlags.Static |
               BindingFlags.DeclaredOnly |
               BindingFlags.Public)
           .Select(c => (Color)c.GetValue(null, null))
           .ToList();
    }
}
