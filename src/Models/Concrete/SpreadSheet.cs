using Microsoft.AspNetCore.Components.Forms;
using System.Collections.Generic;

namespace Simple.ExportToExcel
{
    /// <summary>
    /// Represents a spreadsheet
    /// </summary>
    public class SpreadSheet : ImportDocument
    {
        readonly IList<SheetRow> rows = new List<SheetRow>();

        public SpreadSheet() { }
        public SpreadSheet(IBrowserFile file)
        {
            Size = file.Size;
            FileName = file.Name;
            ContentType = file.ContentType;
        }

        public int RowCount => rows.Count;
        public ICollection<SheetRow> DataRows => rows;
        public SpreadSheet AddRow(object Value)
        {
            rows.Add(new SheetRow(RowCount, Value));

            return this;
        }
        public SpreadSheet AddRow(SheetRow row)
        {
            rows.Add(row);

            return this;
        }
    }
}
