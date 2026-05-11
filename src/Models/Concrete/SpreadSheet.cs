using Microsoft.AspNetCore.Components.Forms;

namespace Simple.ExportToExcel;

/// <summary>
/// Represents a spreadsheet
/// </summary>
public class SpreadSheet : ImportDocument
{
    readonly IList<SheetRow> _rows = new List<SheetRow>();

    /// <summary>
    /// Constructor.
    /// </summary>
    public SpreadSheet() { }

    /// <summary>
    /// Constructor.
    /// </summary>
    public SpreadSheet(IBrowserFile file)
    {
        Size = file.Size;
        FileName = file.Name;
        ContentType = file.ContentType;
    }

    /// <summary>The total number of data rows currently in the spreadsheet.</summary>
    public int RowCount => _rows.Count;

    /// <summary>The collection of all data rows in the spreadsheet.</summary>
    public ICollection<SheetRow> DataRows => _rows;

    /// <summary>
    /// Adds a new row by wrapping <paramref name="Value"/> in a <see cref="SheetRow"/> at the next available index.
    /// </summary>
    /// <param name="Value">The row data to add. Comma-delimited strings are split into columns.</param>
    /// <returns>The current <see cref="SpreadSheet"/> instance for fluent chaining.</returns>
    public SpreadSheet AddRow(object Value)
    {
        _rows.Add(new SheetRow(RowCount, Value));

        return this;
    }

    /// <summary>
    /// Adds a pre-constructed <see cref="SheetRow"/> to the spreadsheet.
    /// </summary>
    /// <param name="row">The row to add.</param>
    /// <returns>The current <see cref="SpreadSheet"/> instance for fluent chaining.</returns>
    public SpreadSheet AddRow(SheetRow row)
    {
        _rows.Add(row);

        return this;
    }
}
