using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace ExportToExcel
{
    public class HeaderBuilder<T>
    {
        private IRow HeaderRow;
        private ISheet ExcelSheet;
        private readonly ICellStyle CellStyle;

        public HeaderBuilder(HeaderStyle HeaderStyle, IWorkbook workBook)
        {
            this.EntityType = typeof(T);
            this.HeaderStyle = HeaderStyle;
            CellStyle = HeaderStyle.GenerateStyleObject(workBook);
        }

        public Type EntityType { get; }
        public HeaderStyle HeaderStyle { get; }

        public ISheet BuildHeaderWithReflection(string FileName, IWorkbook WorkBook)
        {
            if (WorkBook == null)
                throw new NullReferenceException(nameof(WorkBook));

            ExcelSheet = WorkBook.CreateSheet(FileName);
            HeaderRow = ExcelSheet.CreateRow(0);

            var Properties = typeof(T)
                .GetProperties()
                .ToList();

            var ColumnIndex = 0;

            foreach (var item in Properties)
            {
                var propertyType = item.PropertyType;

                if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(IList<>) && propertyType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                {

                }
                else
                {
                    var MemberInfoArray = EntityType.GetMember(item.Name);
                    if (MemberInfoArray != null && MemberInfoArray[0] != null)
                    {
                        var attribute = MemberInfoArray[0].GetCustomAttribute<DisplayAttribute>();
                        if (attribute != null)
                        {
                            CreateCell(ColumnIndex, attribute.Name);
                            ColumnIndex++;
                        }
                    }
                }
            }

            return ExcelSheet;
        }
        private void CreateCell(int Column, string CellValue)
        {
            var HeaderCell = HeaderRow.CreateCell(Column);
            HeaderCell.CellStyle = CellStyle;
            HeaderCell.SetCellValue(CellValue);
        }
    }
}
