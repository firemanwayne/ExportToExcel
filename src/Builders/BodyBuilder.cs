using Simple.ExportToExcel.Models.Concrete;
using Simple.ExportToExcel.Styles;

namespace Simple.ExportToExcel;

/// <summary>
/// Builds the data rows of an Excel spreadsheet for a collection of <typeparamref name="T"/> entities.
/// Uses the column layout established by the paired <see cref="HeaderBuilder{T}"/>.
/// </summary>
/// <typeparam name="T">The entity type whose instances form the spreadsheet rows.</typeparam>
public class BodyBuilder<T>
{
    ISheet _excelSheet;
    readonly HeaderBuilder<T> _header;
    readonly ICellStyle _bodyCellStyle;
    readonly ICellStyle _headerCellStyle;
    readonly IWorkbook _workbook;
    readonly Func<T, bool> _rowStylePredicate;
    readonly ICellStyle _trueRowCellStyle;
    readonly ICellStyle _falseRowCellStyle;
    readonly IList<ConditionalStyle> _conditionalStyles;
    readonly Dictionary<(short colorId, short baseIndex), ICellStyle> _conditionalStyleCache = new();

    /// <summary>
    /// Initializes a new <see cref="BodyBuilder{T}"/> using settings from an <see cref="ExcelDocumentRequest{T}"/>.
    /// </summary>
    /// <param name="request">The document request containing data, workbook, and style configuration.</param>
    /// <param name="header">The header builder that defines the column layout.</param>
    public BodyBuilder(ExcelDocumentRequest<T> request, HeaderBuilder<T> header)
    {
        BodyStyle = request.BodyStyle;
        DataItems = request.ItemsToExport;
        _header = header;
        _workbook = request.Workbook;
        _headerCellStyle = request.HeaderStyle.CellStyle ?? request.HeaderStyle.GenerateStyleObject(request.Workbook);
        _bodyCellStyle = BodyStyle.GenerateStyleObject(request.Workbook);
        _conditionalStyles = request.ConditionalStyles;

        if (request.RowStyleExpression?.EvaluatedExpression != null)
        {
            _rowStylePredicate = request.RowStyleExpression.EvaluatedExpression.Compile();
            _trueRowCellStyle  = request.RowStyleExpression.TrueStyleResult?.GenerateStyleObject(_workbook);
            _falseRowCellStyle = request.RowStyleExpression.FalseStyleResult?.GenerateStyleObject(_workbook);
        }
    }

    /// <summary>
    /// The style applied to all data cells in the body of the spreadsheet.
    /// </summary>
    public BodyStyle BodyStyle { get; }

    /// <summary>
    /// The collection of entities to write as spreadsheet rows.
    /// </summary>
    public IEnumerable<T> DataItems { get; }

    /// <summary>
    /// Populates the data rows on the given sheet using the entities in <see cref="DataItems"/>.
    /// </summary>
    /// <param name="excelSheet">The sheet to write data rows into.</param>
    public void Build(ISheet excelSheet)
    {
        _excelSheet = excelSheet;

        int rowCount = 1;

        foreach (T item in DataItems)
        {
            CreateRow(rowCount, item);
            rowCount++;
        }
    }

    void CreateRow(int rowCount, T entity)
    {
        IRow bodyRow = _excelSheet.CreateRow(rowCount);

        ICellStyle rowCellStyle = _bodyCellStyle;
        if (_rowStylePredicate != null)
        {
            ICellStyle matched = _rowStylePredicate(entity) ? _trueRowCellStyle : _falseRowCellStyle;
            if (matched != null)
                rowCellStyle = matched;
        }

        List<PropertyInfo> properties = _header
            .ColumnProperties
            .Where(e => !e.PropertyType.IsGenericType)
            .ToList();

        int columnCount = 0;

        foreach (PropertyInfo item in properties)
        {
            PropertyInfo parentProperty = entity.GetType().GetProperty(item.Name);
            object propertyValue = parentProperty.GetValue(entity, null) ?? "";
            CreateCell(bodyRow, columnCount, propertyValue, rowCellStyle);
            columnCount++;
        }
    }

    void CreateCell(IRow bodyRow, int column, object value, ICellStyle cellStyle)
    {
        ICell cell = bodyRow.CreateCell(column);
        cell.CellStyle = ResolveConditionalStyle(value, cellStyle);

        switch (value)
        {
            case int v:
                cell.SetCellValue(v);
                break;

            case string v:
                cell.SetCellValue(v);
                break;

            case double v:
                cell.SetCellValue(v);
                break;

            case float v:
                cell.SetCellValue(v);
                break;

            default:
                cell.SetCellValue($"{value}");
                break;
        }
    }

    /// <summary>
    /// Evaluates the <see cref="ConditionalStyle"/> list against the string representation of
    /// <paramref name="value"/>. The first condition whose predicate returns <c>true</c> wins;
    /// its <see cref="ConditionalStyle.TrueColor"/> is applied. The resulting style is cached
    /// per (color, base style) pair. Returns <paramref name="fallback"/> when no condition matches.
    /// </summary>
    ICellStyle ResolveConditionalStyle(object value, ICellStyle fallback)
    {
        if (_conditionalStyles is not { Count: > 0 })
            return fallback;

        string stringValue = value?.ToString() ?? string.Empty;

        foreach (ConditionalStyle condition in _conditionalStyles)
        {
            if (!condition.Condition.Invoke(stringValue))
                continue;

            var cacheKey = (condition.TrueColor.Id, fallback.Index);
            if (!_conditionalStyleCache.TryGetValue(cacheKey, out ICellStyle cached))
            {
                XSSFCellStyle newStyle = (XSSFCellStyle)_workbook.CreateCellStyle();
                newStyle.CloneStyleFrom(fallback);
                newStyle.SetFillForegroundColor(new XSSFColor(condition.TrueColor.IndexedColor.RGB, null));
                newStyle.FillPattern = FillPattern.SolidForeground;
                _conditionalStyleCache[cacheKey] = newStyle;
                cached = newStyle;
            }

            return cached;
        }

        return fallback;
    }

    void CreateBodyRowFromGenericList(T parent, PropertyInfo entity, ISheet excelSheet, int rowCount)
    {
        string propertyName = entity.Name;

        PropertyInfo property = typeof(T).GetProperty(propertyName);
        object propValue = property.GetValue(parent, null);

        List<PropertyInfo> listPropertyList = entity.PropertyType
            .GetGenericArguments()[0]
            .GetProperties()
            .ToList();

        IRow headerColumnRow = excelSheet.CreateRow(rowCount);
        rowCount++;

        for (int h = 0; h < listPropertyList.Count; h++)
        {
            object[] attributes = listPropertyList[h]?.GetCustomAttributes(typeof(DisplayAttribute), false);
            DisplayAttribute displayAttr = attributes?.Length > 0 ? (DisplayAttribute)attributes[0] : null;
            string displayName = displayAttr?.Name ?? listPropertyList[h].Name;

            ICell bodyHeaderCell = headerColumnRow.CreateCell(h);
            bodyHeaderCell.CellStyle = _headerCellStyle;
            bodyHeaderCell.SetCellValue(displayName);
        }

        List<object> listValues = (propValue as IEnumerable<object>)
            .Cast<object>()
            .ToList();

        foreach (object value in listValues)
        {
            IRow bodyRow = excelSheet.CreateRow(rowCount);
            rowCount++;
            System.Type propertyParentInfo = value.GetType();

            for (int p = 0; p < listPropertyList.Count; p++)
            {
                string propName = listPropertyList[p].Name;
                PropertyInfo listPropertyInfo = propertyParentInfo.GetProperty(propName);
                object objectValue = listPropertyInfo.GetValue(value, null);
                ICell bodyCell = bodyRow.CreateCell(p);
                bodyCell.CellStyle = _bodyCellStyle;
                bodyCell.SetCellValue(objectValue?.ToString());
            }
        }
    }
}
