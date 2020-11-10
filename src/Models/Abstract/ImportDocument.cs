namespace ExportToExcel
{
    public abstract class ImportDocument
    {
        public string ContentType { get; set; }
        public long Size { get; set; }
        public string FileName { get; set; }
    }
}
