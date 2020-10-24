using Microsoft.AspNetCore.Components.Forms;
using System.Collections.Generic;

namespace ExportToExcel
{
    public class SpreadSheet : ImportDocument
    {
        private readonly IList<SheetRow> rows = new List<SheetRow>();

        public SpreadSheet() { }
        public SpreadSheet(IBrowserFile file)
        {
            Size = file.Size;
            FileName = file.Name;
            ContentType = file.ContentType;
        }

        public int RowCount { get => rows.Count; }
        public ICollection<SheetRow> DataRows { get => rows; }
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
