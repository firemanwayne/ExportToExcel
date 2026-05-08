namespace Simple.ExportToExcel;

/// <summary>
/// Abstract base class representing a document that can be imported or processed.
/// </summary>
public abstract class ImportDocument
{
    /// <summary>
    /// MIME content type of the document.
    /// </summary>
    public string ContentType { get; set; }

    /// <summary>
    /// Size of the document in bytes.
    /// </summary>
    public long Size { get; set; }

    /// <summary>
    /// Name of the document file, including extension.
    /// </summary>
    public string FileName { get; set; }
}
