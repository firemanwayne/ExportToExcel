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
    ICellStyle _bodyCellStyle;
    ICellStyle _headerCellStyle;

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
        _headerCellStyle = request.HeaderStyle.CellStyle ?? request.HeaderStyle.GenerateStyleObject(request.Workbook);
        _bodyCellStyle = BodyStyle.GenerateStyleObject(request.Workbook);
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

        _bodyCellStyle = BodyStyle.CellStyle;
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

        List<PropertyInfo> properties = _header
            .ColumnProperties
            .Where(e => !e.PropertyType.IsGenericType)
            .ToList();

        int columnCount = 0;

        foreach (PropertyInfo item in properties)
        {
            string propertyName = item.Name;

            PropertyInfo parentProperty = entity
                .GetType()
                .GetProperty(propertyName);

            if (item.IsList())
            {
                CreateBodyRowFromGenericList(entity, item, _excelSheet, rowCount);
            }
            else
            {
                object propertyValue = parentProperty.GetValue(entity, null) ?? "";

                CreateCell(bodyRow, columnCount, propertyValue);

                columnCount++;
            }
        }
    }

    void CreateCell(IRow bodyRow, int column, object value)
    {
        ICell cell = bodyRow.CreateCell(column);
        cell.CellStyle = _bodyCellStyle;

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
