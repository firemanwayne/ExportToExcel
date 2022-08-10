using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Simple.ExportToExcel
{
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
            var RowCount = 1;

            foreach (var item in DataItems)
            {
                CreateRow(RowCount, item);
                RowCount++;
            }
        }

        void CreateRow(int RowCount, T Entity)
        {
            var BodyRow = ExcelSheet.CreateRow(RowCount);

            var Properties = Header
                .ColumnProperties
                .Where(e => !e.PropertyType.IsGenericType)
                .ToList();

            var ColumnCount = 0;

            foreach (var item in Properties)
            {
                var propertyName = item.Name;

                var ParentProperty = Entity
                    .GetType()
                    .GetProperty(propertyName);

                if (item.IsList())
                    CreateBodyRowFromGenericList(Entity, item, ExcelSheet, RowCount);

                else
                {
                    var propertyValue = ParentProperty.GetValue(Entity, null) ?? "";

                    CreateCell(BodyRow, ColumnCount, propertyValue);

                    ColumnCount++;
                }
            }
        }
        static void CreateCell(IRow BodyRow, int Column, object Value)
        {
            var Cell = BodyRow.CreateCell(Column);            

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
            Cell.CellStyle = BodyCellStyle;
        }
        static void CreateBodyRowFromGenericList(T Parent, PropertyInfo Entity, ISheet ExcelSheet, int RowCount)
        {
            var propertyName = Entity.Name;

            var property = typeof(T).GetProperty(propertyName);
            object propValue = property.GetValue(Parent, null);

            var ListPropertyList = Entity.PropertyType
                .GetGenericArguments()[0]
                .GetProperties()
                .ToList();

            IRow HeaderColumnRow1 = ExcelSheet.CreateRow(RowCount);
            RowCount++;

            for (int h = 0; h < ListPropertyList.Count; h++)
            {
                object[] attributes = ListPropertyList[h]?.GetCustomAttributes(typeof(DisplayAttribute), false);
                var Displayname = ((DisplayAttribute)attributes[0]).Name;

                ICell BodyHeaderCell = HeaderColumnRow1.CreateCell(h);
                BodyHeaderCell.CellStyle = HeaderCellStyle;
                BodyHeaderCell.SetCellValue(Displayname);
            }

            var ListValues = (propValue as IEnumerable<object>)
                .Cast<object>()
                .ToList();

            foreach (var value in ListValues)
            {
                IRow BodyRow1 = ExcelSheet.CreateRow(RowCount);
                RowCount++;
                var propertyParentInfo = value.GetType();

                for (int p = 0; p < ListPropertyList.Count; p++)
                {
                    var PropertyName = ListPropertyList[p].Name;
                    var listPropertyInfo = propertyParentInfo.GetProperty(PropertyName);
                    var objectValue = listPropertyInfo.GetValue(value, null);
                    var BodyCell1 = BodyRow1.CreateCell(p);
                    BodyCell1.CellStyle = BodyCellStyle;
                    BodyCell1.SetCellValue(objectValue?.ToString());
                }
            }
        }
    }
}
