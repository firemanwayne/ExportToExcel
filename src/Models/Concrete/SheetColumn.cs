namespace ExportToExcel
{
    public class SheetColumn
    {
        public SheetColumn(int RowIndex, int ColumnIndex, object Value)
        {
            this.RowIndex = RowIndex;
            this.ColumnIndex = ColumnIndex;
            this.Value = Value;
        }
        public int RowIndex { get;}
        public int ColumnIndex { get; }
        public object Value { get; }
    }
}
