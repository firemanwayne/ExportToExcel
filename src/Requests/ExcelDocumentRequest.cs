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
    /// <param name="FileName">Filename of the spreadsheet</param>
    /// <param name="ItemsToExport">Data to be exported</param>
    /// <param name="HeaderStyle">Header row's style</param>
    /// <param name="BodyStyle">Style of the body rows</param>
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

        HeaderStyle.GenerateStyleObject(Workbook);
        BodyStyle.GenerateStyleObject(Workbook);
    }

    /// <summary>
    /// The CLR type of the objects being exported. Can be set for runtime type inspection.
    /// </summary>
    public Type ObjectType { get; set; }

    /// <summary>
    /// Filename of the spreadsheet
    /// </summary>
    public string FileName { get; set; }

    /// <summary>
    /// Data that will be in the spreadsheet
    /// </summary>
    public IEnumerable<T> ItemsToExport { get; set; } = Enumerable.Empty<T>();

    /// <summary>
    /// Style for the header of the spreadsheet
    /// </summary>
    public HeaderStyle HeaderStyle { get; } = new HeaderStyle();

    /// <summary>
    /// Style for the body of the spreadsheet
    /// </summary>
    public BodyStyle BodyStyle { get; } = new BodyStyle();

    /// <summary>
    /// Instance of the Workbook
    /// </summary>
    public IWorkbook Workbook { get; }
}
