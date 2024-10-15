using System;
using System.Linq.Expressions;

namespace Simple.ExportToExcel;

public class RowStyleExpression<T>
{
    public Expression<Func<T, bool>> EvaluatedExpression { get; private set; }

    public RowStyle TrueStyleResult { get; set; }
    public RowStyle FalseStyleResult { get; set; }

    public void FilterByProperty(T entity, string PropertyName, object ValueToCompare, OperationOperatorEnum Operator)
    {
        object Value = entity.GetType().GetProperty(PropertyName).GetValue(entity, null);

        if (Value is int int1 && ValueToCompare is int int2)
        {
            EvaluatedExpression = x => EvaluateIntegers(int1, int2, Operator);
        }
        else if (Value is DateTime time && ValueToCompare is DateTime time1)
        {
            EvaluatedExpression = x => EvaluateDates(time, time1, Operator);
        }
    }

    static bool EvaluateDates(DateTime Date1, DateTime Date2, OperationOperatorEnum Operator)
    {
        if (Operator is OperationOperatorEnum.GreaterThan)
        {
            return Date1 == Date2;
        }
        else if (Operator is OperationOperatorEnum.GreaterThan)
        {
            return Date1 > Date2;
        }
        else if (Operator is OperationOperatorEnum.GreaterThan)
        {
            return Date1 < Date2;
        }

        return false;
    }
    static bool EvaluateIntegers(int Value1, int Value2, OperationOperatorEnum Operator)
    {
        if (Operator is OperationOperatorEnum.GreaterThan)
        {
            return Value1 == Value2;
        }
        else if (Operator is OperationOperatorEnum.GreaterThan)
        {
            return Value1 > Value2;
        }
        else if (Operator is OperationOperatorEnum.GreaterThan)
        {
            return Value1 < Value2;
        }

        return false;
    }
}
