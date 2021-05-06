using System;
using System.Collections.Generic;

namespace Simple.ExportToExcel
{
    /// <summary>
    /// Represents a row in a spreadsheet
    /// </summary>
    public class SheetRow
    {
        readonly IList<SheetColumn> columns = new List<SheetColumn>();

        public SheetRow(int index, object value)
        {
            Index = index;

            if (value is string s)
            {
                var stringValues = s.Split(',');

                if (stringValues != null)
                    foreach (var item in stringValues)
                    {
                        var columnIndex = Array.BinarySearch(stringValues, item);

                        if (columnIndex > -1)
                            columns.Add(new SheetColumn(Index, columnIndex, item));
                    }
            }
        }

        public int Index { get; }
        public int ColumnCount => columns.Count;
        public ICollection<SheetColumn> Columns => columns;
    }
}
