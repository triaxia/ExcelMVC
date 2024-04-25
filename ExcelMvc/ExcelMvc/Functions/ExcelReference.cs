
using ExcelMvc.Runtime;
using ExcelMvc.Views;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Functions
{
    public class ExcelReference
    {
        public string BookName { get; }
        public string SheetName { get; }
        public int RowOffset { get; }
        public int ColumnOffset { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        private string Address { get; }

        public object Value
        {
            get
            {
                return ToRange().Value;
            }
            set
            {
                AsyncActions.Post(_ =>
                {
                    ToRange().Value = value;
                }, null, false);
            }
        }

        internal ExcelReference()
        {
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active workbook.
        /// </summary>
        /// <param name="bookName"></param>
        /// <param name="sheetName"></param>
        /// <param name="rowOffset"></param>
        /// <param name="columnOffset"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        public ExcelReference(string bookName, string sheetName, int rowOffset, int columnOffset, int rowCount, int columnCount)
            : this(GetRange(bookName, sheetName, rowOffset, columnOffset, rowCount, columnCount))
        {
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active workbook.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowOffset"></param>
        /// <param name="columnOffset"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        public ExcelReference(string sheetName, int rowOffset, int columnOffset, int rowCount, int columnCount)
            : this(GetRange(App.Instance.Underlying.ActiveWorkbook.Name, sheetName
                , rowOffset, columnOffset, rowCount, columnCount))
        {
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active worksheet.
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="columnOffset"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        public ExcelReference(int rowOffset, int columnOffset, int rowCount, int columnCount)
            : this(GetRange(App.Instance.Underlying.ActiveWorkbook.Name
                , App.Instance.Underlying.ActiveSheet.Name as string
                , rowOffset, columnOffset, rowCount, columnCount))
        {
        }

        internal ExcelReference(Range range)
        {
            BookName = range.Parent.Parent.Name;
            SheetName = range.Parent.Name;
            RowOffset = range.Row;
            ColumnOffset = range.Column;
            RowCount = range.Rows.Count;
            ColumnCount = range.Columns.Count;
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

        private Range ToRange()
        {
            return GetRange(BookName, SheetName, RowOffset, RowCount, ColumnOffset, ColumnCount);
        }

        private static Range GetRange(string bookName, string sheetName
            , int rowOffset, int columnOffset, int rowCount, int columnCount)
        {
            var sheet = App.Instance.Underlying.Workbooks[bookName].Worksheets[sheetName] as Worksheet;
            var start = sheet.Cells[rowOffset, columnOffset];
            var end = start.Cells[rowOffset + rowCount - 1, columnOffset + columnCount - 1];
            return sheet.Range[start, end] as Range;
        }
    }
}
