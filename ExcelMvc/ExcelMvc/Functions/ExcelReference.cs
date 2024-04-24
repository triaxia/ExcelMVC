
namespace ExcelMvc.Functions
{
    public class ExcelReference
    {
        public string BookName { get; }
        public string SheetName { get; }
        public int FirstRow { get; }
        public int FirstColumn { get; }
        public int LastRow { get; }
        public int LastColumn { get; }
        public string Address { get; }

        public ExcelReference()
        {
        }

        internal ExcelReference(Microsoft.Office.Interop.Excel.Range range)
        {
            BookName = range.Parent.Parent.Name;
            SheetName = range.Parent.Name;
            FirstRow = range.Row;
            FirstColumn = range.Column;
            LastRow = range.Row + range.Rows.Count - 1;
            LastColumn = range.Column + range.Columns.Count - 1;
            Address = range.Address;
        }

        public override string ToString()
        {
            var bn = string.IsNullOrWhiteSpace(BookName) 
                ? string.Empty : $"[{BookName}]";
            var sn = string.IsNullOrWhiteSpace(SheetName)
                ? string.Empty : $"{SheetName}!";
            return $"{bn}{sn}{Address}";
        }
    }
}
