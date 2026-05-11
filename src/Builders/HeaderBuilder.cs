
using Simple.ExportToExcel.Attributes;

namespace Simple.ExportToExcel;

/// <summary>
/// Builds the header row of an Excel spreadsheet for a given entity type <typeparamref name="T"/>.
/// Column names are resolved from <see cref="SpreadSheetAttribute"/> metadata when present,
/// otherwise from <see cref="System.ComponentModel.DataAnnotations.DisplayAttribute"/> or the property name.
/// </summary>
/// <typeparam name="T">The entity type whose properties map to spreadsheet columns.</typeparam>
public class HeaderBuilder<T>
{
    IRow _headerRow;
    ISheet _excelSheet;
    SpreadSheetAttribute _attribute;
    readonly ICellStyle _cellStyle;
    readonly IWorkbook _workBook;
    readonly IList<ICell> _headerCells = [];
    readonly IList<PropertyInfo> _columnProperties = [];

    /// <summary>
    /// Initializes a new <see cref="HeaderBuilder{T}"/> using settings from an <see cref="ExcelDocumentRequest{T}"/>.
    /// </summary>
    /// <param name="request">The document request containing workbook and header style configuration.</param>
    public HeaderBuilder(ExcelDocumentRequest<T> request)
    {
        EntityType = typeof(T);
        _workBook = request.Workbook;
        _cellStyle = request.HeaderStyle.GenerateStyleObject(_workBook);
    }

    /// <summary>
    /// Initializes a new <see cref="HeaderBuilder{T}"/> with an explicit workbook and header style.
    /// </summary>
    /// <param name="headerStyle">The style to apply to header cells.</param>
    /// <param name="workBook">The NPOI workbook instance.</param>
    public HeaderBuilder(HeaderStyle headerStyle, IWorkbook workBook)
    {
        EntityType = typeof(T);
        _workBook = workBook;
        _cellStyle = headerStyle.GenerateStyleObject(workBook);
    }

    /// <summary>
    /// The CLR type of the entity being exported.
    /// </summary>
    public Type EntityType { get; }

    /// <summary>
    /// The ordered list of properties that correspond to generated spreadsheet columns.
    /// </summary>
    public IEnumerable<PropertyInfo> ColumnProperties { get => _columnProperties; }

    /// <summary>
    /// The header cells created during <see cref="Build"/>.
    /// </summary>
    public IEnumerable<ICell> HeaderCells { get => _headerCells; }

    /// <summary>
    /// Creates the worksheet and builds the header row.
    /// </summary>
    /// <param name="fileName">The name of the sheet tab in the workbook.</param>
    /// <returns>The <see cref="ISheet"/> with the header row populated.</returns>
    public ISheet Build(string fileName)
    {
        _excelSheet = _workBook.CreateSheet(fileName);
        _headerRow = _excelSheet.CreateRow(0);

        if (TryGetSpreadSheetMetaData())
        {
            CreateCellsByAttribute();
        }
        else
        {
            CreateCellsByReflection();
        }

        return _excelSheet;
    }

    void AddCell(int column, string cellValue)
    {
        ICell headerCell = _headerRow.CreateCell(column);
        headerCell.CellStyle = _cellStyle;
        headerCell.SetCellValue(cellValue);

        _headerCells.Add(headerCell);
    }

    bool TryGetSpreadSheetMetaData()
    {
        _attribute = typeof(T).GetCustomAttribute<SpreadSheetAttribute>();

        return _attribute != null;
    }

    void CreateCellsByAttribute()
    {
        foreach (SpreadSheetColumnAttribute item in _attribute.Columns)
        {
            _columnProperties.Add(item.Property);

            AddCell(item.ColumnIndex, item.Name);
        }
    }

    void CreateCellsByReflection()
    {
        List<PropertyInfo> properties = EntityType
                .GetProperties()
                .ToList();

        int columnIndex = 0;

        foreach (PropertyInfo item in properties)
        {
            if (item.IsList())
                continue;

            _columnProperties.Add(item);

            MemberInfo[] memberInfoArray = EntityType.GetMember(item.Name);
            if (memberInfoArray != null && memberInfoArray[0] != null)
            {
                DisplayAttribute attribute = memberInfoArray[0].GetCustomAttribute<DisplayAttribute>();
                AddCell(columnIndex, attribute?.Name ?? item.Name);
                columnIndex++;
            }
        }
    }
}
