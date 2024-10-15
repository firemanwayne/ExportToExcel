namespace Simple.ExportToExcel;

using NPOI.SS.UserModel;

using Simple.ExportToExcel.Attributes;

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

public class HeaderBuilder<T>
{
    IRow HeaderRow;
    ISheet ExcelSheet;
    SpreadSheetAttribute Attribute;
    readonly ICellStyle CellStyle;
    readonly IWorkbook workBook;
    readonly IList<ICell> headerCells = [];
    readonly IList<PropertyInfo> columnProperties = [];

    public HeaderBuilder(ExcelDocumentRequest<T> Request)
    {
        EntityType = typeof(T);
        workBook = Request.Workbook;
        CellStyle = Request.HeaderStyle.GenerateStyleObject(workBook);
    }

    public HeaderBuilder(HeaderStyle HeaderStyle, IWorkbook workBook)
    {
        EntityType = typeof(T);
        this.workBook = workBook;
        CellStyle = HeaderStyle.GenerateStyleObject(workBook);
    }

    public Type EntityType { get; }
    public IEnumerable<PropertyInfo> ColumnProperties { get => columnProperties; }
    public IEnumerable<ICell> HeaderCells { get => headerCells; }

    public ISheet Build(string FileName)
    {
        if (workBook == null)
        {
            throw new NullReferenceException(nameof(workBook));
        }

        ExcelSheet = workBook.CreateSheet(FileName);
        HeaderRow = ExcelSheet.CreateRow(0);

        if (TryGetSpreadSheetMetaData())
        {
            CreateCellsByAttribute();
        }
        else
        {
            CreateCellsByReflection();
        }

        return ExcelSheet;
    }

    void AddCell(int Column, string CellValue)
    {
        ICell HeaderCell = HeaderRow.CreateCell(Column);
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
        foreach (SpreadSheetColumnAttribute item in Attribute.Columns)
        {
            columnProperties.Add(item.Property);

            AddCell(item.ColumnIndex, item.Name);
        }
    }
    void CreateCellsByReflection()
    {
        List<PropertyInfo> Properties = EntityType
                .GetProperties()
                .ToList();

        int ColumnIndex = 0;

        foreach (PropertyInfo item in Properties)
        {
            columnProperties.Add(item);

            if (item.IsList())
            {

            }
            else
            {
                MemberInfo[] MemberInfoArray = EntityType.GetMember(item.Name);
                if (MemberInfoArray != null && MemberInfoArray[0] != null)
                {
                    DisplayAttribute attribute = MemberInfoArray[0].GetCustomAttribute<DisplayAttribute>();
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