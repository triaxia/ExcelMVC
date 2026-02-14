using ExcelMvc.Diagnostics;
using ExcelMvc.Rtd;
using ExcelMvc.Runtime;
using ExcelMvc.Windows;
using Function.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelMvc.Functions
{
    public class ExcelFunctionHost : IFunctionHost
    {
        public ExcelFunctionHost()
        {
            XlMarshalExceptionHandler.Failed +=
                    (sender, e) => RaiseFailed(sender, e);
            DelegateFactory.Executing +=
                    (sender, e) => RaiseExecuting(sender, e);
        }

        /// <inheritdoc/>
        public object Application { get { return Views.App.Instance.Underlying; } set { } }
        private Application App => (Application)Views.App.Instance.Underlying;

        /// <inheritdoc/>
        public object ValueMissing => ExcelMissing.Value;

        /// <inheritdoc/>
        public object ValueEmpty => ExcelEmpty.Value;

        /// <inheritdoc/>
        public object ErrorNull => ExcelError.ExcelErrorNull;

        /// <inheritdoc/>
        public object ErrorDiv0 => ExcelError.ExcelErrorDiv0;

        /// <inheritdoc/>
        public object ErrorValue => ExcelError.ExcelErrorValue;

        /// <inheritdoc/>
        public object ErrorRef => ExcelError.ExcelErrorRef;

        /// <inheritdoc/>
        public object ErrorName => ExcelError.ExcelErrorName;

        /// <inheritdoc/>
        public object ErrorNum => ExcelError.ExcelErrorNum;

        /// <inheritdoc/>
        public object ErrorNA => ExcelError.ExcelErrorNA;

        /// <inheritdoc/>
        public object ErrorData => ExcelError.ExcelErrorGettingData;

        /// <inheritdoc/>
        public int RtdThrottleIntervalMilliseconds
        {
            get => App?.RTD.ThrottleInterval ?? 0;
            set
            {
                if (App != null)
                    App.RTD.ThrottleInterval = value;
            }
        }

        /// <inheritdoc/>
        public string StatusBarText
        {
            get { return ((string)App?.StatusBar) ?? ""; }
            set
            {
                if (App != null)
                    App.StatusBar = value;
            }
        }

        /// <inheritdoc/>
        public bool ExecutingEventRaised { get; set; }

        /// <inheritdoc/>
        public Func<Exception, object> ExceptionToFunctionResult
        {
            get { return XlMarshalExceptionHandler.ExceptionToFunctionResult; }
            set { XlMarshalExceptionHandler.ExceptionToFunctionResult = value; }
        }

        /// <inheritdoc/>
        public event EventHandler<MessageEventArgs> Posted;

        /// <inheritdoc/>
        public event EventHandler<RegisteringEventArgs> Registering;

        /// <inheritdoc/>
        public event EventHandler<ErrorEventArgs> Failed;

        /// <inheritdoc/>
        public event EventHandler<ExecutingEventArgs> Executing;

        /// <inheritdoc/>
        public event EventHandler<RtdServerUpdatedEventArgs> RtdUpdated;

        /// <inheritdoc/>
        public IDictionary<object, string> ErrorStrings { get; }
            = new Dictionary<object, string>
        {
            { ExcelError.ExcelErrorNull,"#NULL!" },
            { ExcelError.ExcelErrorDiv0,"#DIV0!" },
            { ExcelError.ExcelErrorValue,"#VALUE!" },
            { ExcelError.ExcelErrorRef,"#REF!" },
            { ExcelError.ExcelErrorName,"#NAME?" },
            { ExcelError.ExcelErrorNum,"#NUM!" },
            { ExcelError.ExcelErrorNA,"#N/A" },
            { ExcelError.ExcelErrorGettingData,"#Data!" }
        };

        /// <inheritdoc/>
        public IDictionary<object, int> ErrorNumbers { get; }
            = new Dictionary<object, int>
        {
            { ExcelError.ExcelErrorNull, -2146826288 },
            { ExcelError.ExcelErrorDiv0, -2146826281 },
            { ExcelError.ExcelErrorValue, -2146826273 },
            { ExcelError.ExcelErrorRef, -2146826265 },
            { ExcelError.ExcelErrorName, -2146826259 },
            { ExcelError.ExcelErrorNum, -2146826252 },
            { ExcelError.ExcelErrorNA, -2146826246 },
            { ExcelError.ExcelErrorGettingData, -2146826245 }
        };

        /// <inheritdoc/>
        public IntPtr GetAsyncHandle(IntPtr handle)
        {
            unsafe
            {
                var p = (XLOPER12*)handle.ToPointer();
                return p->bigdata.data;
            }
        }

        /// <inheritdoc/>
        public void SetAsyncValue(IntPtr handle, object value)
        {
            Run(nameof(XLFunctions.xlAsyncReturn), XLFunctions.xlAsyncReturn, handle, value);
        }

        /// <inheritdoc/>
        public RangeReference GetCallerReference()
        {
            try
            {
                var reference = Run(nameof(XLFunctions.xlfCaller), XLFunctions.xlfCaller) as SheetReference;
                if (reference?.Range == null)
                {
                    // last resort!?...
                    dynamic caller = App?.Caller;
                    var result = caller is Range range ? RangeToReference(range) : null;
                    return result;
                }

                var cell = reference.Range;
                var name = reference.SheetID == IntPtr.Zero ?
                    Run(nameof(XLFunctions.xlSheetNm), XLFunctions.xlSheetNm, cell) as string
                    : Run(nameof(XLFunctions.xlSheetNm), XLFunctions.xlSheetNm, reference) as string;
                var address = Run(nameof(XLFunctions.xlfAddress), XLFunctions.xlfAddress
                    , new object[] {cell.RowFirst + 1, cell.ColumnFirst + 1}) as string;

                return new RangeReference("", name, cell.RowFirst, cell.RowLast
                    , cell.ColumnFirst, cell.ColumnLast, address);

            }
            catch (Exception ex)
            {
                RaiseFailed(this, new ErrorEventArgs(ex));
                return null;
            }
        }

        /// <inheritdoc/>
        public RangeReference GetActiveBookReference(string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            if (App == null) return null;
            var range = GetRange(App.ActiveWorkbook.Name, pageName
                , rowFirst, rowLast, columnFirst, columnLast);
            return range == null ? null : RangeToReference(range);
        }

        /// <inheritdoc/>
        public RangeReference GetActiveSheetReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            if (App == null) return null;
            var range = GetRange(App.ActiveWorkbook.Name
                , App.ActiveSheet.Name
                , rowFirst, rowLast, columnFirst, columnLast);
            return range == null ? null : RangeToReference(range);
        }

        /// <inheritdoc/>
        public object GetRangeValue(RangeReference range)
        {
            return GetRange(range)?.Value;
        }

        /// <inheritdoc/>
        public void SetRangeValue(RangeReference range, object value, bool async)
        {
            var x = GetRange(range);
            if (async)
            {
                AsyncActions.Post(_ => { x.Value = value; }, null);
            }
            else
            {
                x.Value = value;
            }
        }

        /// <inheritdoc/>
        public RangeReference GetReference(string bookName, string pageName
            , int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            var range = GetRange(bookName, pageName
                , rowFirst, rowLast, columnFirst, columnLast);
            return RangeToReference(range);
        }

        /// <inheritdoc/>
        public bool IsInFunctionWizard() => DllImports.IsInFunctionWizard();

        /// <inheritdoc/>
        public void RaiseExecuting(object sender, ExecutingEventArgs args)
        {
            Executing?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public void RaiseFailed(object sender, ErrorEventArgs args)
        {
            Messages.Instance.AddErrorLine(args.GetException());
            Failed?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public void RaisePosted(object sender, MessageEventArgs args)
        {
            Messages.Instance.AddInfoLine(args.Message);
            Posted?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public void RaiseRegistering(object sender, RegisteringEventArgs args)
        {
            Registering?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public object Rtd<TRtdServerImpl>(Func<TRtdServerImpl> implFactory, string server, params string[] args)
            where TRtdServerImpl : class, IRtdServerImpl
        {
            using (var reg = new RtdRegistry(typeof(TRtdServerImpl), () => implFactory?.Invoke()))
            {
                return Rtd(reg.ProgId, server, args);
            }
        }

        /// <inheritdoc/>
        public object Rtd(string progId, string server, params string[] args)
        {
            args = new string[] { progId, server }.Concat(args).ToArray();
            return Run(nameof(XLFunctions.xlfRtd), XLFunctions.xlfRtd, args);
        }

        /// <inheritdoc/>
        public object Run(int function, params object[] args)
        {
            return Run($"{function}", function, args);
        }

        private object Run(string name, int function, params object[] args)
        {
            try
            {
                (var status, var result) = XLCall.Call(function, args);
                if (status != 0)
                {
                    var ex = new ErrorEventArgs(new Exception($"{name} failed (code = {status})."));
                    RaiseFailed(this, ex);
                }
                return result;

            }
            catch(Exception ex)
            {
                RaiseFailed(this, new ErrorEventArgs(ex));
                return ex.Message;
            }
        }

        /// <inheritdoc/>
        public void RegisterFunctions(FunctionDefinitions functions)
        {
            if (Registering != null)
            {
                for (var idx = 0; idx < functions.FunctionCount; idx++)
                {
                    var args = new RegisteringEventArgs(functions.Items[idx]);
                    RaiseRegistering(this, args);
                    functions.Items[idx] = args.Function;
                }
            }

            XLRegistration.Register(functions);
            using (var pFunction = new StructIntPtr<FunctionDefinitions>(ref functions))
            {
                AddIn.RegisterFunctions(pFunction.Ptr);
            }
        }

        private Range GetRange(RangeReference reference)
        {
            try
            {
                if (App == null) return null;
                var sheet = App.Workbooks[reference.BookName]
                    .Worksheets[reference.SheetName] as Worksheet;
                var start = sheet.Cells[reference.RowFirst, reference.ColumnFirst];
                var end = start.Cells[reference.RowLast, reference.ColumnLast];
                return sheet.Range[start, end] as Range;
            }
            catch
            {
                return null;
            }
        }

        private Range GetRange(string bookName, string sheetName
            , int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            try
            {
                if (App == null) return null;
                var sheet = App.Workbooks[bookName]
                    .Worksheets[sheetName] as Worksheet;
                var start = sheet.Cells[rowFirst, columnFirst];
                var end = start.Cells[rowLast, columnLast];
                return sheet.Range[start, end] as Range;
            }
            catch
            {
                return null;
            }
        }

        private static RangeReference RangeToReference(Range range)
        {
            return range == null ? new RangeReference("", "", 0, 0, 0, 0, "")
                : new RangeReference((string)range.Parent.Parent.Name, (string)range.Parent.Name
                    , range.Row, range.Row + range.Rows.Count - 1
                    , range.Column, range.Column + range.Columns.Count - 1
                    , range.Address);
        }

        /// <inheritdoc/>
        public void RaiseRtdUpdated(object sender, RtdServerUpdatedEventArgs args)
        {
            RtdUpdated?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public void Post(Action<object> action, object state)
        {
            AsyncActions.Post(action, state);
        }

        /// <inheritdoc/>
        public string Version => $"{App.Version}.{App.Build}";

        /// <inheritdoc/>
        public bool IsIdeOpen
        {
            get
            {
                try
                {
                    var window = App?.ActiveWorkbook.VBProject.VBE.ActiveWindow;
                    return window != null && window.WindowState != Microsoft.Vbe.Interop.vbext_WindowState.vbext_ws_Minimize;
                }
                catch
                {
                    return false;
                }
            }
        }

        /// <inheritdoc/>
        public string ModuleFileName => AddIn.ModuleFileName;
    }
}
