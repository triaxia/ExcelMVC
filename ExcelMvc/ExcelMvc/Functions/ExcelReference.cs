
using ExcelMvc.Runtime;
using ExcelMvc.Views;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Functions
{
    public class ExcelReference
    {
        public string BookName { get; }
        public string SheetName { get; }
        public int RowFirst { get; }
        public int RowLast { get; }
        public int ColumnFirst { get; }
        public int ColumnLast { get; }
        public string Address { get; }

        public object GetValue() => ToRange().Value;

        public void SetValue(object value, bool async)
        {
            if (async)
            {
                AsyncActions.Post(_ =>
                {
                    ToRange().Value = value;
                }, null, false);
            }
            else
            {
                ToRange().Value = value;
            }
        }

        internal ExcelReference()
        {
        }

        public static ExcelReference GetCaller()
        {
            dynamic caller = App.Instance.Underlying?.Caller;
            return caller is Range range ? new ExcelReference(range)
                : new ExcelReference();
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active workbook.
        /// </summary>
        /// <param name="bookName"></param>
        /// <param name="sheetName"></param>
        /// <param name="rowFirst"></param>
        /// <param name="rowLast"></param>
        /// <param name="columnFirst"></param>
        /// <param name="columnLast"></param>
        public ExcelReference(string bookName, string sheetName, int rowFirst, int rowLast, int columnFirst, int columnLast)
            : this(GetRange(bookName, sheetName, rowFirst, rowLast, columnFirst, columnLast))
        {
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active workbook.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowFirst"></param>
        /// <param name="rowLast"></param>
        /// <param name="columnFirst"></param>
        /// <param name="columnLast"></param>
        public ExcelReference(string sheetName, int rowFirst, int rowLast, int columnFirst, int columnLast)
            : this(GetRange(App.Instance.Underlying.ActiveWorkbook.Name, sheetName
                , rowFirst, rowLast, columnFirst, columnLast))
        {
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active worksheet.
        /// </summary>
        /// <param name="rowFirst"></param>
        /// <param name="rowLast"></param>
        /// <param name="columnFirst"></param>
        /// <param name="columnLast"></param>
        public ExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
            : this(GetRange(App.Instance.Underlying.ActiveWorkbook.Name
                , App.Instance.Underlying.ActiveSheet.Name as string
                , rowFirst, rowLast, columnFirst, columnLast))
        {
        }

        internal ExcelReference(Range range)
        {
            BookName = range.Parent.Parent.Name;
            SheetName = range.Parent.Name;
            RowFirst = range.Row;
            RowLast = RowFirst + range.Rows.Count - 1;
            ColumnFirst = range.Column;
            ColumnLast = RowFirst + range.Columns.Count - 1;
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
            return GetRange(BookName, SheetName, RowFirst, RowLast, ColumnFirst, ColumnLast);
        }

        private static Range GetRange(string bookName, string sheetName
            , int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            var sheet = App.Instance.Underlying.Workbooks[bookName]
                .Worksheets[sheetName] as Worksheet;
            var start = sheet.Cells[rowFirst, columnFirst];
            var end = start.Cells[rowLast, columnLast];
            return sheet.Range[start, end] as Range;
        }
    }
}
