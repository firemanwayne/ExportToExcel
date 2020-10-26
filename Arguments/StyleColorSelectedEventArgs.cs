using ExportToExcel.Styles;
using System;

namespace ExportToExcel
{
    public class StyleColorSelectedEventArgs : EventArgs
    {
        public short ColorIndex { get; }
        public string RGBValue { get; }
        public string ColorName { get; }

        public StyleColorSelectedEventArgs(ExcelColors SelectedColor)
        {
            ColorIndex = SelectedColor.Id;
            RGBValue = SelectedColor.RGBValue;
            ColorName = SelectedColor.Name;
        }
    }
}
