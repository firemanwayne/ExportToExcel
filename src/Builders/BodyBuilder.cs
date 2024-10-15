using NPOI.SS.UserModel;

using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Simple.ExportToExcel;

public class BodyBuilder<T>
{
    ISheet ExcelSheet;
    readonly HeaderBuilder<T> Header;
    static ICellStyle BodyCellStyle;
    static ICellStyle HeaderCellStyle;

    public BodyBuilder(ExcelDocumentRequest<T> Request, HeaderBuilder<T> Header)
    {
        BodyStyle = Request.BodyStyle;
        DataItems = Request.ItemsToExport;
        this.Header = Header;
        HeaderCellStyle = Request.HeaderStyle.CellStyle ?? Request.HeaderStyle.GenerateStyleObject(Request.Workbook);
        BodyCellStyle = BodyStyle.GenerateStyleObject(Request.Workbook);
    }

    public BodyStyle BodyStyle { get; }
    public IEnumerable<T> DataItems { get; }
    public void Build(ISheet ExcelSheet)
    {
        this.ExcelSheet = ExcelSheet;

        BodyCellStyle = BodyStyle.CellStyle;
        int RowCount = 1;

        foreach (T item in DataItems)
        {
            CreateRow(RowCount, item);
            RowCount++;
        }
    }

    void CreateRow(int RowCount, T Entity)
    {
        IRow BodyRow = ExcelSheet.CreateRow(RowCount);

        List<PropertyInfo> Properties = Header
            .ColumnProperties
            .Where(e => !e.PropertyType.IsGenericType)
            .ToList();

        int ColumnCount = 0;

        foreach (PropertyInfo item in Properties)
        {
            string propertyName = item.Name;

            PropertyInfo ParentProperty = Entity
                .GetType()
                .GetProperty(propertyName);

            if (item.IsList())
            {
                CreateBodyRowFromGenericList(Entity, item, ExcelSheet, RowCount);
            }
            else
            {
                object propertyValue = ParentProperty.GetValue(Entity, null) ?? "";

                CreateCell(BodyRow, ColumnCount, propertyValue);

                ColumnCount++;
            }
        }
    }
    static void CreateCell(IRow BodyRow, int Column, object Value)
    {
        ICell Cell = BodyRow.CreateCell(Column);
        Cell.CellStyle = BodyCellStyle;

        switch (Value)
        {
            case int v:
                Cell.SetCellValue(v);
                break;

            case string v:
                Cell.SetCellValue(v);
                break;

            case double v:
                Cell.SetCellValue(v);
                break;

            case float v:
                Cell.SetCellValue(v);
                break;

            default:
                Cell.SetCellValue($"{Value}");
                break;
        }
    }
    static void CreateBodyRowFromGenericList(T Parent, PropertyInfo Entity, ISheet ExcelSheet, int RowCount)
    {
        string propertyName = Entity.Name;

        PropertyInfo property = typeof(T).GetProperty(propertyName);
        object propValue = property.GetValue(Parent, null);

        List<PropertyInfo> ListPropertyList = Entity.PropertyType
            .GetGenericArguments()[0]
            .GetProperties()
            .ToList();

        IRow HeaderColumnRow1 = ExcelSheet.CreateRow(RowCount);
        RowCount++;

        for (int h = 0; h < ListPropertyList.Count; h++)
        {
            object[] attributes = ListPropertyList[h]?.GetCustomAttributes(typeof(DisplayAttribute), false);
            string DisplayName = ((DisplayAttribute)attributes[0]).Name;

            ICell BodyHeaderCell = HeaderColumnRow1.CreateCell(h);
            BodyHeaderCell.CellStyle = HeaderCellStyle;
            BodyHeaderCell.SetCellValue(DisplayName);
        }

        List<object> ListValues = (propValue as IEnumerable<object>)
            .Cast<object>()
            .ToList();

        foreach (object value in ListValues)
        {
            IRow BodyRow1 = ExcelSheet.CreateRow(RowCount);
            RowCount++;
            System.Type propertyParentInfo = value.GetType();

            for (int p = 0; p < ListPropertyList.Count; p++)
            {
                string PropertyName = ListPropertyList[p].Name;
                PropertyInfo listPropertyInfo = propertyParentInfo.GetProperty(PropertyName);
                object objectValue = listPropertyInfo.GetValue(value, null);
                ICell BodyCell1 = BodyRow1.CreateCell(p);
                BodyCell1.CellStyle = BodyCellStyle;
                BodyCell1.SetCellValue(objectValue?.ToString());
            }
        }
    }
}
