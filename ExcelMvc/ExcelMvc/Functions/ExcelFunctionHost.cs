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
using System.Windows.Forms;
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
            FunctionAttributeType = typeof(FunctionAttribute);
            ArgumentAttributeType = typeof(ArgumentAttribute);
        }

        /// <inheritdoc/>
        public object Application { get; set; } = Views.App.Instance.Underlying;
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
        public Type FunctionAttributeType { get; set; }

        /// <inheritdoc/>
        public Type ArgumentAttributeType { get; set; }

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
        public IDictionary<object, string> ErrorMappings { get; }
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
        public string ErrorToString(object value)
        {
            return ErrorMappings.TryGetValue(value, out var mapped) ? mapped : $"{value}";
        }

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
            var xlHandle = XLOPER12.FromObject(handle);
            var xlValue = XLOPER12.FromObject(value);
            try
            {
                using (var p1 = new StructIntPtr<XLOPER12>(ref xlHandle))
                using (var p2 = new StructIntPtr<XLOPER12>(ref xlValue))
                {
                    unsafe
                    {
                        var ptr = AddIn.SetAsyncValue(p1.Ptr, p2.Ptr);
                        var status = (CallStatus*)ptr.ToPointer();
                        var code = status->status;
                        AddIn.FreeCallStatus(ptr);
                        if (code != 0)
                            throw new Exception($"SetAsyncValue failed. (status = {code})");
                    }
                }
            }
            finally
            {
                xlHandle.Dispose();
                xlValue.Dispose();
            }
        }

        /// <inheritdoc/>
        public RangeReference GetCallerReference()
        {
            try
            {
                dynamic caller = App?.Caller;
                return caller is Range range ? RangeToReference(range) : null;
            }
            catch
            {
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
                AsyncActions.Post(_ =>
                {
                    x.Value = value;
                }, null, false);
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
        public object Rtd<TRtdServerImpl>(Func<IRtdServerImpl> implFactory, string server, params string[] args)
            where TRtdServerImpl : IRtdServerImpl
        {
            using (var reg = new RtdRegistry(typeof(IRtdServerImpl), implFactory))
            {
                return Rtd(reg.ProgId, server, args);
            }
        }

        /// <inheritdoc/>
        public object Rtd(string progId, string server, params string[] args)
        {
            var arguments = new string[] { progId, server }
                .Concat(args)
                .Select((x, idx) => new FunctionArgument($"p{idx}", x))
                .ToArray();
            var fArgs = new FunctionArguments(arguments);
            IntPtr ptr = IntPtr.Zero;
            using (var pArgs = new StructIntPtr<FunctionArguments>(ref fArgs))
            {
                ptr = AddIn.CallRtd(pArgs.Ptr);
            }
            unsafe
            {
                var status = (CallStatus*)ptr.ToPointer();
                var code = status->status;
                var obj = status->result == null ? null : status->result->ToObject();
                AddIn.FreeCallStatus(ptr);
                if (code != 0)
                    throw new Exception($"CallRtd failed. (status = {code})");
                return obj;
            }
        }

        /// <inheritdoc/>
        public void RegisterFunctions(FunctionDefinitions functions)
        {
            if (Registering != null)
            {
                for (var idx = 0; idx < functions.Items.Length; idx++)
                {
                    var args = new RegisteringEventArgs(functions.Items[idx]);
                    RaiseRegistering(this, args);
                    functions.Items[idx] = args.Function;
                }
            }

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
            AsyncActions.Post(action, state, false);
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
