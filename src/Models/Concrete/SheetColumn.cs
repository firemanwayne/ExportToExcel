namespace Simple.ExportToExcel;

/// <summary>
/// Represents a single cell (column) within a <see cref="SheetRow"/>.
/// </summary>
public class SheetColumn
{
    /// <summary>
    /// Initializes a new <see cref="SheetColumn"/> with its position and value.
    /// </summary>
    /// <param name="rowIndex">The zero-based row index of the parent row.</param>
    /// <param name="columnIndex">The zero-based column index within the row.</param>
    /// <param name="value">The cell value.</param>
    public SheetColumn(int rowIndex, int columnIndex, object value)
    {
        RowIndex = rowIndex;

        ColumnIndex = columnIndex;

        Value = value;
    }

    /// <summary>
    /// The zero-based index of the row this column belongs to.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// The zero-based column index within the row.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// The value held in this cell.
    /// </summary>
    public object Value { get; }
}
