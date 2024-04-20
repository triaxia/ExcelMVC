
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Functions
{
    public class RangeReference
    {
        public string BookName { get; }
        public string SheetName { get; }
        public int FirstRow { get; }
        public int FirstColumn { get; }
        public int LastRow { get; }
        public int LastColumn { get; }
        
        internal RangeReference(Range range)
        {
            BookName = range.Parent.Parent.Name;
            SheetName = range.Parent.Name;
            FirstRow = range.Row;
            FirstColumn = range.Column;
            LastRow = range.Row + range.Rows.Count;
            LastColumn = range.Column + range.Columns.Count;
        }
    }
}
