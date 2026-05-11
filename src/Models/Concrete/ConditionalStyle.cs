namespace Simple.ExportToExcel.Models.Concrete;

/// <summary>
/// Applies a conditional color to a cell by evaluating a predicate against a string value
/// and returning either a true-color or false-color <see cref="ExcelColors"/> result.
/// </summary>
public class ConditionalStyle
{
    /// <summary>
    /// Initializes a new <see cref="ConditionalStyle"/> with the evaluation parameters.
    /// </summary>
    /// <param name="condition">The predicate invoked with each cell's string value at render time.</param>
    /// <param name="trueColor">The color to apply when <paramref name="condition"/> returns <c>true</c>.</param>
    /// <param name="falseColor">The color associated with a <c>false</c> result (informational; not applied by the pipeline).</param>
    public ConditionalStyle(Func<string, bool> condition, ExcelColors trueColor, ExcelColors falseColor)
    {
        Condition = condition;
        TrueColor = trueColor;
        FalseColor = falseColor;
    }

    /// <summary>The color applied when <see cref="Condition"/> returns <c>true</c>.</summary>
    public ExcelColors TrueColor { get; }

    /// <summary>The color associated with a <c>false</c> result.</summary>
    public ExcelColors FalseColor { get; }

    /// <summary>
    /// The predicate invoked by the pipeline with each cell's string value.
    /// When it returns <c>true</c>, <see cref="TrueColor"/> is applied to the cell.
    /// </summary>
    public Func<string, bool> Condition { get; }
}
