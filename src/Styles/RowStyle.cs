using NPOI.HSSF.Util;

namespace Simple.ExportToExcel;

/// <summary>
/// Defines the cell style for a specific row, supporting border and fill configuration.
/// Used with <see cref="RowStyleExpression{T}"/> for conditional row styling.
/// </summary>
public class RowStyle
{
    /// <summary>
    /// HSSF palette index for the cell foreground fill color. Defaults to white.
    /// </summary>
    public short ForegroundColor { get; set; } = HSSFColor.White.Index;

    /// <summary>
    /// Fill pattern applied to the cell. Defaults to <see cref="FillPattern.SolidForeground"/>.
    /// </summary>
    public FillPattern ForegroundPattern { get; set; } = FillPattern.SolidForeground;

    /// <summary>Top border style. Defaults to <see cref="BorderStyle.Thin"/>.</summary>
    public BorderStyle TopStyle { get; set; } = BorderStyle.Thin;

    /// <summary>Right border style. Defaults to <see cref="BorderStyle.Thin"/>.</summary>
    public BorderStyle RightStyle { get; set; } = BorderStyle.Thin;

    /// <summary>Bottom border style. Defaults to <see cref="BorderStyle.Thin"/>.</summary>
    public BorderStyle BottomStyle { get; set; } = BorderStyle.Thin;

    /// <summary>Left border style. Defaults to <see cref="BorderStyle.Thin"/>.</summary>
    public BorderStyle LeftStyle { get; set; } = BorderStyle.Thin;

    /// <summary>
    /// The generated NPOI <see cref="ICellStyle"/> after <see cref="GenerateStyleObject"/> is called.
    /// </summary>
    public ICellStyle RowCellStyle { get; private set; }

    /// <summary>
    /// Creates and configures an NPOI <see cref="ICellStyle"/> from the current property values.
    /// </summary>
    /// <param name="workBook">The NPOI workbook used to create the cell style.</param>
    /// <returns>The configured <see cref="ICellStyle"/>.</returns>
    public ICellStyle GenerateStyleObject(in IWorkbook workBook)
    {
        RowCellStyle = workBook.CreateCellStyle();
        RowCellStyle.BorderTop = TopStyle;
        RowCellStyle.BorderRight = RightStyle;
        RowCellStyle.BorderBottom = BottomStyle;
        RowCellStyle.BorderLeft = LeftStyle;
        RowCellStyle.FillForegroundColor = ForegroundColor;
        RowCellStyle.FillPattern = ForegroundPattern;

        return RowCellStyle;
    }
}
