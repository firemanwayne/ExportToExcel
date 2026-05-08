using System.Linq.Expressions;

namespace Simple.ExportToExcel;

/// <summary>
/// Represents a conditional style expression that evaluates a property of <typeparamref name="T"/>
/// against a value using an <see cref="OperationOperatorEnum"/> operator and returns the appropriate
/// <see cref="RowStyle"/> based on whether the condition is true or false.
/// </summary>
/// <typeparam name="T">The entity type whose properties are evaluated.</typeparam>
public class RowStyleExpression<T>
{
    /// <summary>
    /// The compiled LINQ expression representing the evaluated condition.
    /// Populated after calling <see cref="FilterByProperty"/>.
    /// </summary>
    public Expression<Func<T, bool>> EvaluatedExpression { get; private set; }

    /// <summary>
    /// The <see cref="RowStyle"/> to apply when the condition evaluates to <c>true</c>.
    /// </summary>
    public RowStyle TrueStyleResult { get; set; }

    /// <summary>
    /// The <see cref="RowStyle"/> to apply when the condition evaluates to <c>false</c>.
    /// </summary>
    public RowStyle FalseStyleResult { get; set; }

    /// <summary>
    /// Evaluates the specified property of <paramref name="entity"/> against <paramref name="valueToCompare"/>
    /// using the given <paramref name="filterOperator"/> and stores the resulting expression.
    /// </summary>
    /// <param name="entity">The entity instance to evaluate.</param>
    /// <param name="propertyName">Name of the property on <typeparamref name="T"/> to compare.</param>
    /// <param name="valueToCompare">The value to compare the property against.</param>
    /// <param name="filterOperator">The comparison operator to apply.</param>
    public void FilterByProperty(T entity, string propertyName, object valueToCompare, OperationOperatorEnum filterOperator)
    {
        object value = entity.GetType().GetProperty(propertyName).GetValue(entity, null);

        if (value is int int1 && valueToCompare is int int2)
        {
            EvaluatedExpression = x => EvaluateIntegers(int1, int2, filterOperator);
        }
        else if (value is DateTime time && valueToCompare is DateTime time1)
        {
            EvaluatedExpression = x => EvaluateDates(time, time1, filterOperator);
        }
    }

    static bool EvaluateDates(DateTime date1, DateTime date2, OperationOperatorEnum filterOperator)
    {
        if (filterOperator is OperationOperatorEnum.GreaterThan)
        {
            return date1 == date2;
        }
        else if (filterOperator is OperationOperatorEnum.GreaterThan)
        {
            return date1 > date2;
        }
        else if (filterOperator is OperationOperatorEnum.GreaterThan)
        {
            return date1 < date2;
        }

        return false;
    }

    static bool EvaluateIntegers(int value1, int value2, OperationOperatorEnum filterOperator)
    {
        if (filterOperator is OperationOperatorEnum.GreaterThan)
        {
            return value1 == value2;
        }
        else if (filterOperator is OperationOperatorEnum.GreaterThan)
        {
            return value1 > value2;
        }
        else if (filterOperator is OperationOperatorEnum.GreaterThan)
        {
            return value1 < value2;
        }

        return false;
    }
}
