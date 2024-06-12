using ExcelDna.Integration;
using ExcelMvc.Functions;
using Function.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDnaAddIn
{
    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    public class ExcelDnaHost : IFunctionHost
    {
        public object Application { get; set; } = ExcelDnaUtil.Application;
        private Application App => (Application)ExcelDnaUtil.Application;

        public object ValueMissing => ExcelMissing.Value;

        public object ValueEmpty => ExcelEmpty.Value;

        public object ErrorNull => ExcelError.ExcelErrorNull;

        public object ErrorDiv0 => ExcelError.ExcelErrorDiv0;

        public object ErrorValue => ExcelError.ExcelErrorValue;

        public object ErrorRef => ExcelError.ExcelErrorRef;

        public object ErrorName => ExcelError.ExcelErrorName;

        public object ErrorNA => ExcelError.ExcelErrorNA;

        public object ErrorData => ExcelError.ExcelErrorGettingData;

        public int RtdThrottleIntervalMilliseconds
        {
            get => App.RTD.ThrottleInterval;
            set => App.RTD.ThrottleInterval = value;
        }
        public string StatusBarText
        {
            get => App.StatusBar;
            set => App.StatusBar = value;
        }
        public bool ExecutingEventRaised { get; set; } = false;

        public Func<Exception, object> ExceptionToFunctionResult { get; set; }
            = e => $"{e.Message}";

        public Type FunctionAttributeType { get; set; } = typeof(FunctionAttribute);
        public Type ArgumentAttributeType { get; set; } = typeof(ArgumentAttribute);

        public string Version => $"{App.Version}.{App.Build}";

        public bool IsIdeOpen
        {
            get
            {
                var window = App.ActiveWorkbook.VBProject.VBE.ActiveWindow;
                return window != null && window.WindowState != Microsoft.Vbe.Interop.vbext_WindowState.vbext_ws_Minimize;
            }
        }

        public event EventHandler<RtdServerUpdatedEventArgs> RtdUpdated;
        public event EventHandler<MessageEventArgs> Posted;
        public event EventHandler<RegisteringEventArgs> Registering;
        public event EventHandler<ErrorEventArgs> Failed;
        public event EventHandler<ExecutingEventArgs> Executing;

        public static Dictionary<object, string> Mappings = new Dictionary<object, string>
            {
                { ExcelError.ExcelErrorNull,"#NULL!" },
                { ExcelError.ExcelErrorDiv0,"#DIV0!" },
                { ExcelError.ExcelErrorValue,"#VALUE!" },
                { ExcelError.ExcelErrorRef,"#REF!" },
                { ExcelError.ExcelErrorName,"#NAME?" },
                { ExcelError.ExcelErrorNum,"#NUM!" },
                { ExcelError.ExcelErrorNA,"#N/A" },
                { ExcelError.ExcelErrorGettingData,"#Data!" },
                { ExcelMissing.Value, $"{ExcelMissing.Value}" },
                { ExcelEmpty.Value, $"{ExcelEmpty.Value}" }
            };

        public string ErrorToString(object value)
        {
            return Mappings.TryGetValue(value, out var mapped) ? mapped : $"{value}";
        }

        public RangeReference GetCallerReference()
        {
            dynamic caller = App?.Caller;
            return caller is Range range ? RangeToReference(range)
                : RangeToReference(null);
        }

        public RangeReference GetActiveBookReference(string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            var range = GetRange(App.ActiveWorkbook.Name, pageName
                , rowFirst, rowLast, columnFirst, columnLast);
            return RangeToReference(range);
        }

        public RangeReference GetActivePageReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            var range = GetRange(App.ActiveWorkbook.Name
                , App.ActiveSheet.Name
                , rowFirst, rowLast, columnFirst, columnLast);
            return RangeToReference(range);
        }

        public object GetRangeValue(RangeReference range)
        {
            return GetRange(range)?.Value;
        }

        public RangeReference GetReference(string bookName, string pageName
            , int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            var range = GetRange(bookName, pageName
                , rowFirst, rowLast, columnFirst, columnLast);
            return RangeToReference(range);
        }

        public bool IsInFunctionWizard()
            => ExcelDnaUtil.IsInFunctionWizard();

        public void RaiseExecuting(object sender, ExecutingEventArgs args)
        {
            Executing?.Invoke(sender, args);
        }

        public void RaiseFailed(object sender, ErrorEventArgs args)
        {
            Failed?.Invoke(sender, args);
        }

        public void RaisePosted(object sender, MessageEventArgs args)
        {
            Posted?.Invoke(sender, args);
        }

        public void RaiseRegistering(object sender, RegisteringEventArgs args)
        {
            Registering?.Invoke(sender, args);
        }

        public void RaiseRtdUpdated(object sender, RtdServerUpdatedEventArgs args)
        {
            RtdUpdated?.Invoke(sender, args);
        }

        public void RegisterFunctions(FunctionDefinitions functions)
        {
        }

        public void SetRangeValue(RangeReference range, object value, bool async)
        {
            var x = GetRange(range);
            if (async)
            {
                Post(r =>
                {
                    ((Range)r).Value = value;
                }, range);
            }
            else
            {
                x.Value = value;
            }
        }

        public object Rtd<TRtdServerImpl>(Func<IRtdServerImpl> implFactory, string server, params string[] args)
            where TRtdServerImpl : IRtdServerImpl
        {
            return XlCall.RTD(typeof(TRtdServerImpl).FullName, "", args);
        }

        public object Rtd(string progId, string server, params string[] args)
        {
            return XlCall.RTD(progId, "", args);
        }

        public void SetAsyncResult(IntPtr handle, object result)
        {
            throw new NotImplementedException();
        }

        public void Post(Action<object> action, object state)
        {
            ExcelAsyncUtil.QueueAsMacro(x => action(x), state);
        }

        private Range GetRange(RangeReference reference)
        {
            var sheet = App.Workbooks[reference.BookName]
                .Worksheets[reference.PageName] as Worksheet;
            var start = sheet.Cells[reference.RowFirst, reference.ColumnFirst];
            var end = start.Cells[reference.RowLast, reference.ColumnLast];
            return sheet.Range[start, end] as Range;
        }

        private Range GetRange(string bookName, string sheetName
            , int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            var sheet = App.Workbooks[bookName]
                .Worksheets[sheetName] as Worksheet;
            var start = sheet.Cells[rowFirst, columnFirst];
            var end = start.Cells[rowLast, columnLast];
            return sheet.Range[start, end] as Range;
        }

        private static RangeReference RangeToReference(Range range)
        {
            return range == null ? new RangeReference("", "", 0, 0, 0, 0, "")
                : new RangeReference((string)range.Parent.Parent.Name, (string)range.Parent.Name
                    , range.Row, range.Row + range.Rows.Count - 1
                    , range.Column, range.Column + range.Columns.Count - 1
                    , range.Address);
        }

        public IntPtr GetAsyncHandle(IntPtr handle)
        {
            throw new NotImplementedException();
        }
    }
}
