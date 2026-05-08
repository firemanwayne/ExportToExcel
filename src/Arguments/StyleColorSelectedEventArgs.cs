namespace Simple.ExportToExcel;

/// <summary>
/// Event arguments raised when the user selects a color in the style picker UI.
/// Carries the palette index, CSS RGB string, and display name of the chosen color.
/// </summary>
public class StyleColorSelectedEventArgs : EventArgs
{
    /// <summary>The NPOI palette index of the selected color.</summary>
    public short ColorIndex { get; }

    /// <summary>The CSS RGB string of the selected color (e.g. <c>rgb(255, 0, 0)</c>).</summary>
    public string RGBValue { get; }

    /// <summary>The human-readable display name of the selected color.</summary>
    public string ColorName { get; }

    /// <summary>
    /// Initializes a new <see cref="StyleColorSelectedEventArgs"/> from the given <see cref="ExcelColors"/> entry.
    /// </summary>
    /// <param name="selectedColor">The color selected by the user.</param>
    public StyleColorSelectedEventArgs(ExcelColors selectedColor)
    {
        ColorIndex = selectedColor.Id;
        RGBValue = selectedColor.RGBValue;
        ColorName = selectedColor.Name;
    }
}
