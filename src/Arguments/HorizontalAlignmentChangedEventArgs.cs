namespace Simple.ExportToExcel;

/// <summary>
/// Event arguments raised when the user changes the horizontal alignment setting in the style picker UI.
/// </summary>
public class HorizontalAlignmentChangedEventArgs : EventArgs
{
    /// <summary>The horizontal alignment value selected by the user.</summary>
    public HorizontalAlignment SelectedAlignment { get; }

    /// <summary>
    /// Initializes a new <see cref="HorizontalAlignmentChangedEventArgs"/> with the chosen alignment.
    /// </summary>
    /// <param name="args">The selected <see cref="HorizontalAlignment"/> value.</param>
    public HorizontalAlignmentChangedEventArgs(HorizontalAlignment args)
    {
        SelectedAlignment = args;
    }
}
