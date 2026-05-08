namespace Simple.ExportToExcel;

/// <summary>
/// Represents a file extension paired with its associated MIME type value.
/// </summary>
public struct Extension
{
    /// <summary>
    /// The file extension, including the leading dot (e.g. <c>.xlsx</c>).
    /// </summary>
    public string Key { get; set; }

    /// <summary>
    /// The MIME type string associated with this extension (e.g. <c>application/vnd.openxmlformats-officedocument.spreadsheetml.sheet</c>).
    /// </summary>
    public string Value { get; set; }

    /// <summary>
    /// Initializes a new <see cref="Extension"/> with the given file extension and MIME type.
    /// </summary>
    /// <param name="key">The file extension including leading dot.</param>
    /// <param name="value">The corresponding MIME type string.</param>
    public Extension(string key, string value)
    {
        Key = key;
        Value = value;
    }
}