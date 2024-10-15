using Simple.ExportToExcel.Styles;

using System;

namespace Simple.ExportToExcel.Models.Concrete;

public class ConditionalStyle
{
    public ConditionalStyle(Func<string, bool> condition, string value, ExcelColors trueColor, ExcelColors falseColor)
    {
        Condition = condition;
        Value = value;
        TrueColor = trueColor;
        FalseColor = falseColor;
    }

    public ExcelColors TrueColor { get; }
    public ExcelColors FalseColor { get; }
    public string Value { get; }
    public Func<string, bool> Condition { get; }

    public ExcelColors Evaluate()
    {
        return Condition.Invoke(Value) ? TrueColor : FalseColor;
    }
}
