using Simple.ExportToExcel.Attributes;
using System.ComponentModel.DataAnnotations;

namespace Simple.ExportToExcel.UnitTests;

public class DisplayModel
{
    [Display(Name = "First Name")]
    public string FirstName { get; set; } = string.Empty;

    [Display(Name = "Age")]
    public int Age { get; set; }

    [Display(Name = "Score")]
    public double Score { get; set; }
}

public class PlainModel
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}

[SpreadSheet("AttributeModel", typeof(AttributeModel))]
public class AttributeModel
{
    [SpreadSheetColumn("Column B", 1)]
    public string B { get; set; } = string.Empty;

    [SpreadSheetColumn("Column A", 0)]
    public string A { get; set; } = string.Empty;
}

public class FloatModel
{
    public float Ratio { get; set; }
    public string Label { get; set; } = string.Empty;
}

public class MixedTypeModel
{
    public string Text { get; set; } = string.Empty;
    public bool Flag { get; set; }
}

public class ListPropertyModel
{
    public string Name { get; set; } = string.Empty;
    public IList<string> Items { get; set; } = new List<string>();
    public IEnumerable<int> Numbers { get; set; } = Enumerable.Empty<int>();
}

public class DateModel
{
    public DateTime Date { get; set; }
    public int Count { get; set; }
}
