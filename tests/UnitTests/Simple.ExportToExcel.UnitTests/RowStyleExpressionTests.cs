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
        var expr  = new RowStyleExpression<PlainModel>();
        var style = new RowStyle();

        expr.TrueStyleResult = style;

        Assert.AreSame(style, expr.TrueStyleResult);
    }

    [TestMethod]
    public void FalseStyleResult_CanBeSetAndRead()
    {
        var expr  = new RowStyleExpression<PlainModel>();
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
        var expr = new RowStyleExpression<PlainModel>();

        expr.FilterByProperty(nameof(PlainModel.Value), 10, OperationOperatorEnum.Equals);

        Assert.IsNotNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_DateTimeProperty_SetsEvaluatedExpression()
    {
        var expr = new RowStyleExpression<DateModel>();

        expr.FilterByProperty(nameof(DateModel.Date), DateTime.UtcNow, OperationOperatorEnum.Equals);

        Assert.IsNotNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_StringProperty_EqualsOperator_SetsEvaluatedExpression()
    {
        var expr = new RowStyleExpression<PlainModel>();

        expr.FilterByProperty(nameof(PlainModel.Name), "hello", OperationOperatorEnum.Equals);

        Assert.IsNotNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_StringProperty_UnsupportedOperator_LeavesExpressionNull()
    {
        var expr = new RowStyleExpression<PlainModel>();

        expr.FilterByProperty(nameof(PlainModel.Name), "hello", OperationOperatorEnum.GreaterThan);

        Assert.IsNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_MismatchedTypes_LeavesExpressionNull()
    {
        var expr = new RowStyleExpression<PlainModel>();

        expr.FilterByProperty(nameof(PlainModel.Value), DateTime.UtcNow, OperationOperatorEnum.Equals);

        Assert.IsNull(expr.EvaluatedExpression);
    }

    [TestMethod]
    public void FilterByProperty_CalledTwice_OverwritesExpression()
    {
        var expr = new RowStyleExpression<PlainModel>();

        expr.FilterByProperty(nameof(PlainModel.Value), 1, OperationOperatorEnum.GreaterThan);
        var first = expr.EvaluatedExpression;

        expr.FilterByProperty(nameof(PlainModel.Value), 2, OperationOperatorEnum.GreaterThan);

        Assert.AreNotSame(first, expr.EvaluatedExpression);
    }

    // -----------------------------------------------------------------------
    // EvaluateIntegers
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluateIntegers_EqualsOperator_EqualValues_ReturnsTrue()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 5 };

        expr.FilterByProperty(nameof(PlainModel.Value), 5, OperationOperatorEnum.Equals);

        Assert.IsTrue(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateIntegers_EqualsOperator_UnequalValues_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 5 };

        expr.FilterByProperty(nameof(PlainModel.Value), 99, OperationOperatorEnum.Equals);

        Assert.IsFalse(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateIntegers_GreaterThan_EntityValueGreater_ReturnsTrue()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 10 };

        expr.FilterByProperty(nameof(PlainModel.Value), 3, OperationOperatorEnum.GreaterThan);

        Assert.IsTrue(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateIntegers_GreaterThan_EqualValues_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 7 };

        expr.FilterByProperty(nameof(PlainModel.Value), 7, OperationOperatorEnum.GreaterThan);

        Assert.IsFalse(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateIntegers_LessThan_EntityValueLess_ReturnsTrue()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 2 };

        expr.FilterByProperty(nameof(PlainModel.Value), 10, OperationOperatorEnum.LessThan);

        Assert.IsTrue(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateIntegers_LessThan_EntityValueGreater_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Value = 10 };

        expr.FilterByProperty(nameof(PlainModel.Value), 2, OperationOperatorEnum.LessThan);

        Assert.IsFalse(expr.EvaluatedExpression.Compile()(entity));
    }

    // -----------------------------------------------------------------------
    // EvaluateDates
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluateDates_EqualsOperator_SameDates_ReturnsTrue()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var date   = new DateTime(2024, 1, 1);
        var entity = new DateModel { Date = date };

        expr.FilterByProperty(nameof(DateModel.Date), date, OperationOperatorEnum.Equals);

        Assert.IsTrue(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateDates_EqualsOperator_DifferentDates_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var entity = new DateModel { Date = new DateTime(2024, 1, 1) };

        expr.FilterByProperty(nameof(DateModel.Date), new DateTime(2025, 1, 1), OperationOperatorEnum.Equals);

        Assert.IsFalse(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateDates_GreaterThan_EntityDateGreater_ReturnsTrue()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var entity = new DateModel { Date = new DateTime(2024, 12, 31) };

        expr.FilterByProperty(nameof(DateModel.Date), new DateTime(2024, 1, 1), OperationOperatorEnum.GreaterThan);

        Assert.IsTrue(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateDates_GreaterThan_SameDates_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var date   = new DateTime(2024, 6, 15);
        var entity = new DateModel { Date = date };

        expr.FilterByProperty(nameof(DateModel.Date), date, OperationOperatorEnum.GreaterThan);

        Assert.IsFalse(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateDates_LessThan_EntityDateLess_ReturnsTrue()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var entity = new DateModel { Date = new DateTime(2020, 1, 1) };

        expr.FilterByProperty(nameof(DateModel.Date), new DateTime(2025, 1, 1), OperationOperatorEnum.LessThan);

        Assert.IsTrue(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateDates_LessThan_EntityDateGreater_ReturnsFalse()
    {
        var expr   = new RowStyleExpression<DateModel>();
        var entity = new DateModel { Date = new DateTime(2025, 6, 1) };

        expr.FilterByProperty(nameof(DateModel.Date), new DateTime(2020, 1, 1), OperationOperatorEnum.LessThan);

        Assert.IsFalse(expr.EvaluatedExpression.Compile()(entity));
    }

    // -----------------------------------------------------------------------
    // String evaluation
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluateString_EqualsOperator_MatchingValue_ReturnsTrue()
    {
        var expr   = new RowStyleExpression<PlainModel>();
        var entity = new PlainModel { Name = "Active" };

        expr.FilterByProperty(nameof(PlainModel.Name), "Active", OperationOperatorEnum.Equals);

        Assert.IsTrue(expr.EvaluatedExpression.Compile()(entity));
    }

    [TestMethod]
    public void EvaluateString_EqualsOperator_NonMatchingValue_ReturnsFalse()
    {
        var expr      = new RowStyleExpression<PlainModel>();
        var entity    = new PlainModel { Name = "Inactive" };

        expr.FilterByProperty(nameof(PlainModel.Name), "Active", OperationOperatorEnum.Equals);

        Assert.IsFalse(expr.EvaluatedExpression.Compile()(entity));
    }

    // -----------------------------------------------------------------------
    // Expression reads from the queried entity, not a captured seed value
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluatedExpression_ReadsFromQueryEntity_NotConstant()
    {
        var expr      = new RowStyleExpression<PlainModel>();
        var different = new PlainModel { Value = 99 };

        expr.FilterByProperty(nameof(PlainModel.Value), 5, OperationOperatorEnum.Equals);

        // Expression is x => x.Value == 5; evaluating against Value=99 → false
        Assert.IsFalse(expr.EvaluatedExpression.Compile()(different));
    }

    // -----------------------------------------------------------------------
    // LINQ integration — filters per row value
    // -----------------------------------------------------------------------

    [TestMethod]
    public void EvaluatedExpression_CanBeUsedWithLinqWhere_FiltersCorrectly()
    {
        var expr = new RowStyleExpression<PlainModel>();
        var data = new[]
        {
            new PlainModel { Value = 5 },
            new PlainModel { Value = 5 },
            new PlainModel { Value = 9 },
        };

        expr.FilterByProperty(nameof(PlainModel.Value), 5, OperationOperatorEnum.Equals);
        var matches = data.AsQueryable().Where(expr.EvaluatedExpression).ToList();

        Assert.AreEqual(2, matches.Count);
    }

    [TestMethod]
    public void EvaluatedExpression_CanBeUsedWithLinqWhere_NoMatchesReturnsEmpty()
    {
        var expr = new RowStyleExpression<PlainModel>();
        var data = new[]
        {
            new PlainModel { Value = 5 },
            new PlainModel { Value = 5 },
            new PlainModel { Value = 9 },
        };

        expr.FilterByProperty(nameof(PlainModel.Value), 99, OperationOperatorEnum.Equals);
        var matches = data.AsQueryable().Where(expr.EvaluatedExpression).ToList();

        Assert.AreEqual(0, matches.Count);
    }

    [TestMethod]
    public void EvaluatedExpression_GreaterThan_FiltersRowsCorrectly()
    {
        var expr = new RowStyleExpression<PlainModel>();
        var data = new[]
        {
            new PlainModel { Value = 3 },
            new PlainModel { Value = 7 },
            new PlainModel { Value = 12 },
        };

        // x => x.Value > 5 → matches 7 and 12
        expr.FilterByProperty(nameof(PlainModel.Value), 5, OperationOperatorEnum.GreaterThan);
        var matches = data.AsQueryable().Where(expr.EvaluatedExpression).ToList();

        Assert.AreEqual(2, matches.Count);
    }
}
