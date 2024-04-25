
using ExcelMvc.Runtime;
using ExcelMvc.Views;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Functions
{
    public class ExcelReference
    {
        public string BookName { get; }
        public string SheetName { get; }
        public int Row { get; }
        public int Column { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        private string Address { get; }

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
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        public ExcelReference(string bookName, string sheetName, int row, int column, int rowCount, int columnCount)
            : this(GetRange(bookName, sheetName, row, column, rowCount, columnCount))
        {
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active workbook.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        public ExcelReference(string sheetName, int row, int column, int rowCount, int columnCount)
            : this(GetRange(App.Instance.Underlying.ActiveWorkbook.Name, sheetName
                , row, column, rowCount, columnCount))
        {
        }

        /// <summary>
        /// Initializes a new instance of <see cref="ExcelReference"/> on the active worksheet.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        public ExcelReference(int row, int column, int rowCount, int columnCount)
            : this(GetRange(App.Instance.Underlying.ActiveWorkbook.Name
                , App.Instance.Underlying.ActiveSheet.Name as string
                , row, column, rowCount, columnCount))
        {
        }

        internal ExcelReference(Range range)
        {
            BookName = range.Parent.Parent.Name;
            SheetName = range.Parent.Name;
            Row = range.Row;
            Column = range.Column;
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
            return GetRange(BookName, SheetName, Row, Column, RowCount, ColumnCount);
        }

        private static Range GetRange(string bookName, string sheetName
            , int row, int column, int rowCount, int columnCount)
        {
            var sheet = App.Instance.Underlying.Workbooks[bookName]
                .Worksheets[sheetName] as Worksheet;
            var start = sheet.Cells[row, column];
            var end = start.Cells[row + rowCount - 1, column + columnCount - 1];
            return sheet.Range[start, end] as Range;
        }
    }
}
