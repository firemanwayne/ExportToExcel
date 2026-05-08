namespace Simple.ExportToExcel.Attributes;

/// <summary>
/// Marks a property as a spreadsheet column with an explicit header name and column position.
/// Used in conjunction with <see cref="SpreadSheetAttribute"/> on the containing class.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class SpreadSheetColumnAttribute : Attribute
{
    /// <summary>
    /// Initializes a new <see cref="SpreadSheetColumnAttribute"/> with the column header name and position.
    /// </summary>
    /// <param name="name">The text to display in the column header cell.</param>
    /// <param name="columnIndex">The zero-based index of this column in the spreadsheet.</param>
    public SpreadSheetColumnAttribute(string name, int columnIndex) : base()
    {
        Name = name;
        ColumnIndex = columnIndex;
    }

    /// <summary>The text displayed in the column header cell.</summary>
    public string Name { get; }

    /// <summary>The zero-based column index in the spreadsheet.</summary>
    public int ColumnIndex { get; }

    /// <summary>The <see cref="PropertyInfo"/> of the decorated property. Set during attribute initialization.</summary>
    public PropertyInfo Property { get; private set; }

    /// <summary>
    /// Assigns the reflected <see cref="PropertyInfo"/> to this attribute. Called by <see cref="SpreadSheetAttribute"/>.
    /// </summary>
    /// <param name="p">The property that carries this attribute.</param>
    internal void SetProperty(PropertyInfo p) => Property = p;
}
