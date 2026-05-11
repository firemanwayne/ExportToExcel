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
    /// Builds a LINQ expression that evaluates the specified property of each <typeparamref name="T"/>
    /// instance against <paramref name="valueToCompare"/> using the given <paramref name="filterOperator"/>.
    /// Supported property types: <see cref="int"/>, <see cref="long"/>, <see cref="double"/>,
    /// <see cref="float"/>, <see cref="decimal"/>, <see cref="DateTime"/>, <see cref="string"/>,
    /// <see cref="bool"/>. String and bool only support <see cref="OperationOperatorEnum.Equals"/>.
    /// </summary>
    /// <param name="propertyName">Name of the property on <typeparamref name="T"/> to compare.</param>
    /// <param name="valueToCompare">The constant value to compare the property against.</param>
    /// <param name="filterOperator">The comparison operator to apply.</param>
    public void FilterByProperty(string propertyName, object valueToCompare, OperationOperatorEnum filterOperator)
    {
        PropertyInfo prop = typeof(T).GetProperty(propertyName);
        if (prop == null) return;

        ParameterExpression parameter = Expression.Parameter(typeof(T), "x");
        MemberExpression member = Expression.Property(parameter, prop);

        Expression body = BuildComparison(member, prop.PropertyType, valueToCompare, filterOperator);
        if (body != null)
            EvaluatedExpression = Expression.Lambda<Func<T, bool>>(body, parameter);
    }

    static Expression BuildComparison(MemberExpression member, Type propType, object valueToCompare, OperationOperatorEnum op)
    {
        if (valueToCompare == null || valueToCompare.GetType() != propType)
            return null;

        ConstantExpression constant = Expression.Constant(valueToCompare, propType);

        if (propType == typeof(string) || propType == typeof(bool))
            return op == OperationOperatorEnum.Equals ? Expression.Equal(member, constant) : null;

        if (propType == typeof(int)     || propType == typeof(long)    ||
            propType == typeof(double)  || propType == typeof(float)   ||
            propType == typeof(decimal) || propType == typeof(DateTime))
        {
            return op switch
            {
                OperationOperatorEnum.Equals      => Expression.Equal(member, constant),
                OperationOperatorEnum.GreaterThan => Expression.GreaterThan(member, constant),
                OperationOperatorEnum.LessThan    => Expression.LessThan(member, constant),
                _ => null
            };
        }

        return null;
    }
}
