namespace Simple.ExportToExcel;

/// <summary>
/// Request object used to encapsulate parameters to generate response
/// </summary>
/// <typeparam name="T">Type of data to be exported</typeparam>
public class ExcelDocumentRequest<T>
{
    /// <summary>
    /// Creates a request object that's used to export data into a spreadsheet
    /// </summary>
    /// <param name="fileName">Filename of the spreadsheet</param>
    /// <param name="itemsToExport">Data to be exported</param>
    /// <param name="headerStyle">Header row's style</param>
    /// <param name="bodyStyle">Style of the body rows</param>
    public ExcelDocumentRequest(string fileName, IEnumerable<T> itemsToExport, HeaderStyle headerStyle = null, BodyStyle bodyStyle = null)
    {
        Workbook = new XSSFWorkbook();

        ItemsToExport = itemsToExport ?? throw new ArgumentNullException($"{nameof(itemsToExport)}: you provided nothing to export");
        HeaderStyle = headerStyle ?? new HeaderStyle();
        BodyStyle = bodyStyle ?? new BodyStyle();


        if (MimeMapping.IsExtensionMissing(fileName))
        {
            FileName = fileName += ".xlsx";
        }
        else
        {
            FileName = fileName;
        }

    }

    /// <summary>
    /// Filename of the spreadsheet
    /// </summary>
    public string FileName { get; }

    /// <summary>
    /// Data that will be in the spreadsheet
    /// </summary>
    public IEnumerable<T> ItemsToExport { get; set; }

    /// <summary>
    /// Style for the header of the spreadsheet
    /// </summary>
    public HeaderStyle HeaderStyle { get; }

    /// <summary>
    /// Style for the body of the spreadsheet
    /// </summary>
    public BodyStyle BodyStyle { get; }

    /// <summary>
    /// Optional conditional row style expression. When set, each data row is evaluated
    /// and styled with <see cref="RowStyleExpression{T}.TrueStyleResult"/> or
    /// <see cref="RowStyleExpression{T}.FalseStyleResult"/> accordingly.
    /// </summary>
    public RowStyleExpression<T> RowStyleExpression { get; set; }

    /// <summary>
    /// Optional list of cell-level conditional styles. For each cell the first condition
    /// whose predicate matches the cell's string value is applied.
    /// </summary>
    public IList<ConditionalStyle> ConditionalStyles { get; set; } = [];

    /// <summary>
    /// Instance of the Workbook
    /// </summary>
    public IWorkbook Workbook { get; }
}
