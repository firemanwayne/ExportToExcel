namespace Simple.ExportToExcel;

/// <summary>
/// Exception thrown when a file name is provided without a file extension.
/// </summary>
internal class FileNameMissingExtensionException : Exception
{
    /// <summary>
    /// Initializes a new instance of <see cref="FileNameMissingExtensionException"/> with a default message.
    /// </summary>
    public FileNameMissingExtensionException() : base("The file name provided is missing the extension")
    { }
}