namespace ExcelMvc.Functions
{
    /// <summary>
    /// Represents a range on the host.
    /// </summary>
    public class RangeReference
    {
        /// <summary>
        /// The book name.
        /// </summary>
        public string BookName { get; }

        /// <summary>
        /// The page name.
        /// </summary>
        public string PageName { get; }

        /// <summary>
        /// The first row index.
        /// </summary>
        public int RowFirst { get; }

        /// <summary>
        /// The last row index.
        /// </summary>
        public int RowLast { get; }

        /// <summary>
        /// The first column index.
        /// </summary>
        public int ColumnFirst { get; }

        /// <summary>
        /// The last column index.
        /// </summary>
        public int ColumnLast { get; }

        /// <summary>
        /// The string representation of the range, without <see cref="BookName"/> and <see cref="PageName"/>.
        /// </summary>
        public string Address { get; }

        /// <summary>
        /// Initialises a new instance of <see cref="RangeReference"/>
        /// </summary>
        /// <param name="bookName"></param>
        /// <param name="pageName"></param>
        /// <param name="rowFirst"></param>
        /// <param name="rowLast"></param>
        /// <param name="columnFirst"></param>
        /// <param name="columnLast"></param>
        /// <param name="address"></param>
        public RangeReference(string bookName, string pageName, int rowFirst, int rowLast
            , int columnFirst, int columnLast, string address)
        {
            BookName = bookName;
            PageName = pageName;
            RowFirst = rowFirst;
            RowLast = rowLast;
            ColumnFirst = columnFirst;
            ColumnLast = columnLast;
            Address = address;
        }

        /// <summary>
        /// <inheritdoc cref="object.ToString"/>
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            var bn = string.IsNullOrWhiteSpace(BookName)
                ? string.Empty : $"[{BookName}]";
            var sn = string.IsNullOrWhiteSpace(PageName)
                ? string.Empty : $"{PageName}!";
            return $"{bn}{sn}{Address}";
        }
    }
}
