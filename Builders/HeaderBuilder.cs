using ExportToExcel.Attributes;
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
        IRow HeaderRow;
        ISheet ExcelSheet;
        SpreadSheetAttribute Attribute;
        readonly ICellStyle CellStyle;
        readonly IWorkbook workBook;
        readonly IList<ICell> headerCells = new List<ICell>();
        readonly IList<PropertyInfo> columnProperties = new List<PropertyInfo>();

        public HeaderBuilder(ExcelDocumentRequest<T> Request)
        {
            EntityType = typeof(T);
            workBook = Request.Workbook;
            HeaderStyle = Request.HeaderStyle;
            CellStyle = HeaderStyle.GenerateStyleObject(workBook);
        }

        public HeaderBuilder(HeaderStyle HeaderStyle, IWorkbook workBook)
        {
            EntityType = typeof(T);
            this.workBook = workBook;
            this.HeaderStyle = HeaderStyle;
            CellStyle = HeaderStyle.GenerateStyleObject(workBook);
        }

        public Type EntityType { get; }
        public HeaderStyle HeaderStyle { get; }
        public IEnumerable<PropertyInfo> ColumnProperties { get => columnProperties; }
        public IEnumerable<ICell> HeaderCells { get => headerCells; }

        public ISheet Build(string FileName)
        {
            if (workBook == null)
                throw new NullReferenceException(nameof(workBook));

            ExcelSheet = workBook.CreateSheet(FileName);
            HeaderRow = ExcelSheet.CreateRow(0);

            if (TryGetSpreadSheetMetaData())            
                CreateCellsByAttribute();   
            
            else            
                CreateCellsByReflection();           

            return ExcelSheet;
        }

        void AddCell(int Column, string CellValue)
        {
            var HeaderCell = HeaderRow.CreateCell(Column);
            HeaderCell.CellStyle = CellStyle;
            HeaderCell.SetCellValue(CellValue);

            headerCells.Add(HeaderCell);
        }
        bool TryGetSpreadSheetMetaData()
        {
            Attribute = typeof(T).GetCustomAttribute<SpreadSheetAttribute>();

            return Attribute != null;
        }
        void CreateCellsByAttribute()
        {
            foreach (var item in Attribute.Columns)
            {
                columnProperties.Add(item.Property);

                AddCell(item.ColumnIndex, item.Name);
            }
        }
        void CreateCellsByReflection()
        {
            var Properties = EntityType
                    .GetProperties()
                    .ToList();

            var ColumnIndex = 0;

            foreach (var item in Properties)
            {
                columnProperties.Add(item);

                if (item.IsList())
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
                            AddCell(ColumnIndex, attribute.Name);
                            ColumnIndex++;
                        }
                    }
                }
            }
        }
    }
}