namespace Simple.ExportToExcel;

/// <summary>
/// Orchestrates the assembly of an Excel spreadsheet by coordinating
/// <see cref="HeaderBuilder{T}"/> and <see cref="BodyBuilder{T}"/>.
/// </summary>
public static class ExcelBuilder
{

    /// <summary>
    /// Builds the Excel spreadsheet by first creating the header row then populating the data rows.
    /// </summary>
    /// <typeparam name="T">The type of entity being exported.</typeparam>
    /// <param name="fileName">The sheet tab name used when creating the worksheet.</param>
    /// <param name="body">The body builder that writes data rows.</param>
    /// <param name="header">The header builder that creates the header row and worksheet.</param>
    public static void Build<T>(string fileName, BodyBuilder<T> body, HeaderBuilder<T> header)
    {
        ISheet sheet = header.Build(fileName);

        body.Build(sheet);
    }
}
