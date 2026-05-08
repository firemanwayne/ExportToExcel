namespace Simple.ExportToExcel;

/// <summary>
/// Comparison operators used when evaluating conditional row style expressions.
/// </summary>
public enum OperationOperatorEnum
{
    /// <summary>Checks that two values are equal.</summary>
    Equals,

    /// <summary>Checks that the left value is greater than the right value.</summary>
    GreaterThan,

    /// <summary>Checks that the left value is less than the right value.</summary>
    LessThan
}
