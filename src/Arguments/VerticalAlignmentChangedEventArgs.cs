namespace Simple.ExportToExcel;

/// <summary>
/// Event arguments raised when the user changes the vertical alignment setting in the style picker UI.
/// </summary>
public class VerticalAlignmentChangedEventArgs : EventArgs
{
    /// <summary>The vertical alignment value selected by the user.</summary>
    public VerticalAlignment SelectedAlignment { get; }

    /// <summary>
    /// Initializes a new <see cref="VerticalAlignmentChangedEventArgs"/> with the chosen alignment.
    /// </summary>
    /// <param name="args">The selected <see cref="VerticalAlignment"/> value.</param>
    public VerticalAlignmentChangedEventArgs(VerticalAlignment args)
    {
        SelectedAlignment = args;
    }
}
