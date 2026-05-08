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
    /// <param name="condition">The predicate to evaluate against <paramref name="value"/>.</param>
    /// <param name="value">The string value passed to <paramref name="condition"/>.</param>
    /// <param name="trueColor">The color to return when the condition is <c>true</c>.</param>
    /// <param name="falseColor">The color to return when the condition is <c>false</c>.</param>
    public ConditionalStyle(Func<string, bool> condition, string value, ExcelColors trueColor, ExcelColors falseColor)
    {
        Condition = condition;
        Value = value;
        TrueColor = trueColor;
        FalseColor = falseColor;
    }

    /// <summary>The color applied when <see cref="Condition"/> returns <c>true</c>.</summary>
    public ExcelColors TrueColor { get; }

    /// <summary>The color applied when <see cref="Condition"/> returns <c>false</c>.</summary>
    public ExcelColors FalseColor { get; }

    /// <summary>The string value evaluated by <see cref="Condition"/>.</summary>
    public string Value { get; }

    /// <summary>The predicate that determines which color is returned by <see cref="Evaluate"/>.</summary>
    public Func<string, bool> Condition { get; }

    /// <summary>
    /// Invokes <see cref="Condition"/> with <see cref="Value"/> and returns
    /// <see cref="TrueColor"/> or <see cref="FalseColor"/> accordingly.
    /// </summary>
    /// <returns>The <see cref="ExcelColors"/> result of the condition evaluation.</returns>
    public ExcelColors Evaluate()
    {
        return Condition.Invoke(Value) ? TrueColor : FalseColor;
    }
}
