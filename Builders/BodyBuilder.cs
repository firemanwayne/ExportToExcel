using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace ExportToExcel
{
    public class BodyBuilder<T>
    {
        private ISheet ExcelSheet;
        private static ICellStyle BodyCellStyle;
        private static ICellStyle HeaderCellStyle;

        public BodyBuilder(IEnumerable<T> DataItems, BodyStyle BodyStyle, IWorkbook workBook, HeaderStyle HeaderStyle)
        {
            this.BodyStyle = BodyStyle;
            this.DataItems = DataItems;

            HeaderCellStyle = HeaderStyle.HeaderCellStyle ?? HeaderStyle.GenerateStyleObject(workBook);
            BodyCellStyle = BodyStyle.GenerateStyleObject(workBook);
        }

        public BodyStyle BodyStyle { get; }
        public IEnumerable<T> DataItems { get; }

        public void BuildBody(ISheet ExcelSheet)
        {
            this.ExcelSheet = ExcelSheet;

            BodyCellStyle = BodyStyle.BodyCellStyle;
            var RowCount = 1;

            foreach (var item in DataItems)
            {
                CreateRow(RowCount, item);
                RowCount++;
            }
        }

        private void CreateRow(int RowCount, T Entity)
        {
            var BodyRow = ExcelSheet.CreateRow(RowCount);

            var Properties = Entity.GetType()
                .GetProperties().Where(e => !e.PropertyType.IsGenericType)
                .ToList();

            var ColumnCount = 0;

            foreach (var item in Properties)
            {
                var propertyName = item.Name;
                var ParentProperty = Entity.GetType().GetProperty(propertyName);

                if (item.PropertyType.IsGenericType && item.PropertyType.GetGenericTypeDefinition() == typeof(IList<>) && item.PropertyType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                {
                    CreateBodyRowFromGenericList(Entity, item, ExcelSheet, RowCount);
                }
                else
                {
                    object propertyValue = ParentProperty.GetValue(Entity, null) ?? "";

                    CreateCell(BodyRow, ColumnCount, $"{propertyValue}");

                    ColumnCount++;
                }
            }
        }

        private void CreateCell(IRow BodyRow, int Column, string Value)
        {
            var Cell = BodyRow.CreateCell(Column);
            Cell.CellStyle = BodyCellStyle;
            Cell.SetCellValue(Value);
        }

        private static void CreateBodyRowFromGenericList(T Parent, PropertyInfo Entity, ISheet ExcelSheet, int RowCount)
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

            for (int h = 0; h < ListPropertyList.Count(); h++)
            {
                object[] attributes = ListPropertyList[h]?.GetCustomAttributes(typeof(DisplayAttribute), false);
                string Displayname = ((DisplayAttribute)attributes[0]).Name;


                ICell BodyHeaderCell = HeaderColumnRow1.CreateCell(h);
                BodyHeaderCell.CellStyle = HeaderCellStyle;
                BodyHeaderCell.SetCellValue(Displayname);
            }

            var ListValues = (propValue as IEnumerable<object>)
                .Cast<object>()
                .ToList();

            foreach (object value in ListValues)
            {
                IRow BodyRow1 = ExcelSheet.CreateRow(RowCount);
                RowCount++;
                Type propertyParentInfo = value.GetType();
                for (int p = 0; p < ListPropertyList.Count(); p++)
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
}
