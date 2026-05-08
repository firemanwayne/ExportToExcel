using System.Linq.Expressions;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class RowStyleExpressionTests
{
    // -----------------------------------------------------------------------
    // Initial state
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluatedExpression_IsNullBeforeFilterByPropertyCalled()
    {
        var expr = new RowStyleExpression<PlainModel>();

        Assert.IsNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void TrueStyleResult_CanBeSetAndRead()
    {
        var expr = new RowStyleExpression<PlainModel>();
        var style = new RowStyle();

        expr.TrueStyleResult = style;

        Assert.AreSame(style, expr.TrueStyleResult);
    }

    [TestMethod]
    public void FalseStyleResult_CanBeSetAndRead()
    {
        var expr = new RowStyleExpression<PlainModel>();
        var style = new RowStyle();

        expr.FalseStyleResult = style;

        Assert.AreSame(style, expr.FalseStyleResult);
    }

    // -----------------------------------------------------------------------
    // FilterByProperty — expression assignment
    // -----------------------------------------------------------------------

    [TestMethod]
    public void FilterByProperty_IntProperty_SetsEvaluatedExpression()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 10 };

        expr.FilterByProperty(entity, nameof(PlainModel.Value), 10, OperationOperatorEnum.Equals);

        Assert.IsNotNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_DateTimeProperty_SetsEvaluatedExpression()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var now    = DateTime.UtcNow;
        var entity = new DateModel { Date = now };

        expr.FilterByProperty(entity, nameof(DateModel.Date), now, OperationOperatorEnum.Equals);

        Assert.IsNotNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_StringProperty_LeavesExpressionNull()
    {
        // FilterByProperty only handles int and DateTime; string is ignored.
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Name = "hello" };

        expr.FilterByProperty(entity, nameof(PlainModel.Name), "hello", OperationOperatorEnum.Equals);

        Assert.IsNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_MismatchedTypes_LeavesExpressionNull()
    {
        // int property compared against a DateTime — no branch matches.
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 5 };

        expr.FilterByProperty(entity, nameof(PlainModel.Value), DateTime.UtcNow, OperationOperatorEnum.Equals);

        Assert.IsNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_CalledTwice_OverwritesExpression()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 1 };

        expr.FilterByProperty(entity, nameof(PlainModel.Value), 1, OperationOperatorEnum.GreaterThan);
        var first = expr.EvaluatedExpression;

        expr.FilterByProperty(entity, nameof(PlainModel.Value), 2, OperationOperatorEnum.GreaterThan);

        Assert.AreNotSame(first, expr.EvaluatedExpression);
    }

    // -----------------------------------------------------------------------
    // EvaluateIntegers — actual (buggy) behaviour
    //
    // All three if-branches check OperationOperatorEnum.GreaterThan, so:
    //   • Equals    → no branch matches → false
    //   • GreaterThan → first branch hits → returns value1 == value2
    //   • LessThan  → no branch matches → false
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluateIntegers_EqualsOperator_ReturnsFalse()
    {
        // Bug: Equals operator falls through all GreaterThan checks → false.
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 5 };

        expr.FilterByProperty(entity, nameof(PlainModel.Value), 5, OperationOperatorEnum.Equals);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateIntegers_GreaterThan_EqualValues_ReturnsTrue()
    {
        // Bug: GreaterThan branch returns value1 == value2, so equal values → true.
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 7 };

        expr.FilterByProperty(entity, nameof(PlainModel.Value), 7, OperationOperatorEnum.GreaterThan);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateIntegers_GreaterThan_UnequalValues_ReturnsFalse()
    {
        // Bug: GreaterThan branch returns value1 == value2, so unequal values → false.
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 10 };

        expr.FilterByProperty(entity, nameof(PlainModel.Value), 3, OperationOperatorEnum.GreaterThan);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateIntegers_LessThanOperator_ReturnsFalse()
    {
        // Bug: LessThan falls through all GreaterThan checks → false.
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 2 };

        expr.FilterByProperty(entity, nameof(PlainModel.Value), 10, OperationOperatorEnum.LessThan);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsFalse(result);
    }

    // -----------------------------------------------------------------------
    // EvaluateDates — same bug, same actual behaviour
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluateDates_EqualsOperator_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var now    = new DateTime(2024, 1, 1);
        var entity = new DateModel { Date = now };

        expr.FilterByProperty(entity, nameof(DateModel.Date), now, OperationOperatorEnum.Equals);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateDates_GreaterThan_SameDates_ReturnsTrue()
    {
        // Bug: GreaterThan branch returns date1 == date2, so same dates → true.
        var expr   = new RowStyleExpression<DateModel>();
        var date   = new DateTime(2024, 6, 15);
        var entity = new DateModel { Date = date };

        expr.FilterByProperty(entity, nameof(DateModel.Date), date, OperationOperatorEnum.GreaterThan);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsTrue(result);
    }

    [TestMethod]
    public void EvaluateDates_GreaterThan_DifferentDates_ReturnsFalse()
    {
        // Bug: GreaterThan branch returns date1 == date2, so different dates → false.
        var expr   = new RowStyleExpression<DateModel>();
        var entity = new DateModel { Date = new DateTime(2024, 12, 31) };

        expr.FilterByProperty(entity, nameof(DateModel.Date), new DateTime(2024, 1, 1), OperationOperatorEnum.GreaterThan);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsFalse(result);
    }

    [TestMethod]
    public void EvaluateDates_LessThanOperator_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var entity = new DateModel { Date = new DateTime(2020, 1, 1) };

        expr.FilterByProperty(entity, nameof(DateModel.Date), new DateTime(2025, 1, 1), OperationOperatorEnum.LessThan);
        bool result = expr.EvaluatedExpression.Compile()(entity);

        Assert.IsFalse(result);
    }

    // -----------------------------------------------------------------------
    // Expression is usable with LINQ
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluatedExpression_CanBeUsedWithLinqWhere()
    {
        var expr = new RowStyleExpression<PlainModel>();
        var data = new[]
        {
            new PlainModel { Value = 5 },
            new PlainModel { Value = 5 },
            new PlainModel { Value = 9 },
        };

        // GreaterThan with equal values returns true (see bug notes above).
        expr.FilterByProperty(data[0], nameof(PlainModel.Value), 5, OperationOperatorEnum.GreaterThan);
        var matches = data.AsQueryable().Where(expr.EvaluatedExpression).ToList();

        // All three rows match because the expression captures value1==value2 (5==5=true) as a constant.
        Assert.AreEqual(3, matches.Count);
    }
}
