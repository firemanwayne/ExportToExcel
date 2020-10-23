using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Linq.Expressions;

namespace ExportToExcel
{
    public class RowStyle
    {
        public short ForegroundColor { get; set; } = HSSFColor.White.Index;
        public FillPattern ForegroundPattern { get; set; } = FillPattern.SolidForeground;
        public BorderStyle TopStyle { get; set; } = BorderStyle.Thin;
        public BorderStyle RightStyle { get; set; } = BorderStyle.Thin;
        public BorderStyle BottomStyle { get; set; } = BorderStyle.Thin;
        public BorderStyle LeftStyle { get; set; } = BorderStyle.Thin;

        public ICellStyle RowCellStyle { get; private set; }

        public ICellStyle GenerateStyleObject(in IWorkbook WorkBook)
        {
            RowCellStyle = WorkBook.CreateCellStyle();
            RowCellStyle.BorderTop = TopStyle;
            RowCellStyle.BorderRight = RightStyle;
            RowCellStyle.BorderBottom = BottomStyle;
            RowCellStyle.BorderLeft = LeftStyle;
            RowCellStyle.FillForegroundColor = ForegroundColor;
            RowCellStyle.FillPattern = ForegroundPattern;

            return RowCellStyle;
        }
    }
    public class RowStyleExpression<T>
    {
        public Expression<Func<T, bool>> EvaluatedExpression { get; private set; }

        public RowStyle TrueStyleResult { get; set; }
        public RowStyle FalseStyleResult { get; set; }

        public void FilterByProperty(T entity, string PropertyName, object ValueToCompare, OperationOperatorEnum Operator)
        {
            var Value = entity.GetType().GetProperty(PropertyName).GetValue(entity, null);

            if (Value is int int1 && ValueToCompare is int int2)
                EvaluatedExpression = x => EvaluateInts(int1, int2, Operator);

            else if (Value is DateTime time && ValueToCompare is DateTime time1)
                EvaluatedExpression = x => EvaluateDates(time, time1, Operator);
        }

        private static bool EvaluateDates(DateTime Date1, DateTime Date2, OperationOperatorEnum Operator)
        {
            if (Operator == OperationOperatorEnum.GreaterThan)
                return Date1 == Date2;

            else if (Operator == OperationOperatorEnum.GreaterThan)
                return Date1 > Date2;

            else if (Operator == OperationOperatorEnum.GreaterThan)
                return Date1 < Date2;

            return false;
        }
        private static bool EvaluateInts(int Value1, int Value2, OperationOperatorEnum Operator)
        {
            if (Operator == OperationOperatorEnum.GreaterThan)
                return Value1 == Value2;

            else if (Operator == OperationOperatorEnum.GreaterThan)
                return Value1 > Value2;

            else if (Operator == OperationOperatorEnum.GreaterThan)
                return Value1 < Value2;

            return false;
        }
    }
    public enum OperationOperatorEnum { Equals, GreaterThan, LessThan }
}
