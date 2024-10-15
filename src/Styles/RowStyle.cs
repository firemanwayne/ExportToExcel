using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace Simple.ExportToExcel;

public class RowStyle
{
    public short ForegroundColor { get; set; } = HSSFColor.White.Index;
    public FillPattern ForegroundPattern { get; set; } = FillPattern.SolidForeground;
    public BorderStyle TopStyle { get; set; } = BorderStyle.Thin;
    public BorderStyle RightStyle { get; set; } = BorderStyle.Thin;
    public BorderStyle BottomStyle { get; set; } = BorderStyle.Thin;
    public BorderStyle LeftStyle { get; set; } = BorderStyle.Thin;

    public ICellStyle RowCellStyle { get; private set; }

    public ICellStyle GenerateStyleObject(in IWorkbook WorkBook)
    {
        RowCellStyle = WorkBook.CreateCellStyle();
        RowCellStyle.BorderTop = TopStyle;
        RowCellStyle.BorderRight = RightStyle;
        RowCellStyle.BorderBottom = BottomStyle;
        RowCellStyle.BorderLeft = LeftStyle;
        RowCellStyle.FillForegroundColor = ForegroundColor;
        RowCellStyle.FillPattern = ForegroundPattern;

        return RowCellStyle;
    }
}
