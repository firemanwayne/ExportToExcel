namespace Simple.ExportToExcel;

/// <summary>
/// Represents a row in a spreadsheet
/// </summary>
public class SheetRow
{
    readonly IList<SheetColumn> _columns = new List<SheetColumn>();

    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="index"></param>
    /// <param name="value"></param>
    public SheetRow(int index, object value)
    {
        Index = index;

        if (value is string s)
        {
            string[] stringValues = s.Split(',');

            if (stringValues != null)
            {
                foreach (string item in stringValues)
                {
                    int columnIndex = Array.IndexOf(stringValues, item);

                    if (columnIndex > -1)
                    {
                        _columns.Add(new SheetColumn(Index, columnIndex, item));
                    }
                }
            }
        }
    }

    /// <summary>The zero-based index of this row within the spreadsheet.</summary>
    public int Index { get; }

    /// <summary>The number of columns in this row.</summary>
    public int ColumnCount => _columns.Count;

    /// <summary>The collection of columns (cells) that make up this row.</summary>
    public ICollection<SheetColumn> Columns => _columns;
}
