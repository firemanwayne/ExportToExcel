namespace Simple.ExportToExcel.Attributes;

/// <summary>
/// Marks a class as an Excel spreadsheet data source and specifies the sheet name.
/// Properties decorated with <see cref="SpreadSheetColumnAttribute"/> are discovered
/// automatically when the attribute is applied.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class SpreadSheetAttribute : Attribute
{
    private readonly IList<SpreadSheetColumnAttribute> _columns = new List<SpreadSheetColumnAttribute>();

    /// <summary>
    /// Initializes a new <see cref="SpreadSheetAttribute"/> for the given name and type,
    /// then reflects over <paramref name="type"/> to collect all <see cref="SpreadSheetColumnAttribute"/> definitions.
    /// </summary>
    /// <param name="name">The name of the spreadsheet sheet tab.</param>
    /// <param name="type">The CLR type whose properties define the column layout.</param>
    public SpreadSheetAttribute(string name, Type type) : base()
    {
        Name = name;
        Type = type;

        GetColumnIndexes();
    }

    /// <summary>The name of the spreadsheet sheet tab.</summary>
    public string Name { get; }

    /// <summary>The CLR type used to discover column definitions via reflection.</summary>
    public Type Type { get; }

    /// <summary>
    /// The ordered collection of columns defined on <see cref="Type"/>, sorted by <see cref="SpreadSheetColumnAttribute.ColumnIndex"/>.
    /// </summary>
    public IEnumerable<SpreadSheetColumnAttribute> Columns { get => _columns.OrderBy(c => c.ColumnIndex); }

    /// <summary>
    /// Reflects over <see cref="Type"/> to find all properties decorated with
    /// <see cref="SpreadSheetColumnAttribute"/> and adds them to the internal columns list.
    /// </summary>
    void GetColumnIndexes()
    {
        PropertyInfo[] props = Type.GetProperties();
        foreach (PropertyInfo p in props)
        {
            SpreadSheetColumnAttribute columnAttr = p.GetCustomAttribute<SpreadSheetColumnAttribute>();
            if (columnAttr != null)
            {
                columnAttr.SetProperty(p);

                _columns.Add(columnAttr);
            }
        }
    }
}
