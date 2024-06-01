using ExcelMvc.Diagnostics;
using ExcelMvc.Rtd;
using ExcelMvc.Views;
using ExcelMvc.Windows;
using Function.Interfaces;
using System;
using System.IO;
using System.Linq;

namespace ExcelMvc.Functions
{
    public class ExcelFunctionHost : IHost
    {
        public ExcelFunctionHost()
        {
            XlMarshalExceptionHandler.Failed +=
                    (sender, e) => RaiseFailed(sender, e);
            DelegateFactory.Executing +=
                    (sender, e) => RaiseExecuting(sender, e);
        }

        /// <inheritdoc/>
        public object Underlying => App.Instance.Underlying;

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
        public object ErrorNA => ExcelError.ExcelErrorNA;

        /// <inheritdoc/>
        public object ErrorData => ExcelError.ExcelErrorGettingData;

        /// <inheritdoc/>
        public int RTDThrottleIntervalMilliseconds
        {
            get => App.Instance.Underlying?.RTD.ThrottleInterval ?? 0;
            set
            {
                if (App.Instance.Underlying != null)
                    App.Instance.Underlying.RTD.ThrottleInterval = value;
            }
        }
        /// <inheritdoc/>
        public string StatusBarText
        {
            get { return ((string)App.Instance.Underlying?.StatusBar) ?? ""; }
            set
            {
                if (App.Instance.Underlying != null)
                    App.Instance.Underlying.StatusBar = value;
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
        public string ErrorToString(object value)
        {
            return ExcelErrorMappings.Mappings.TryGetValue(value, out var mapped) ? mapped : $"{value}";
        }

        /// <inheritdoc/>
        public RangeReference GetActiveBookReference(string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc/>
        public RangeReference GetActivePageReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            throw new NotImplementedException();
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
        public RangeReference GetCallerReference()
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc/>
        public object GetRangeValue(RangeReference range)
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc/>
        public RangeReference GetReference(string bookName, string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            throw new NotImplementedException();
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
        public object RTD<TRtdServerImpl>(Func<IRtdServerImpl> implFactory, string arg0, params string[] args)
            where TRtdServerImpl : IRtdServerImpl
        {
            using (var reg = new RtdRegistry(typeof(TRtdServerImpl), implFactory))
            {
                return RTD(reg.ProgId, arg0, args);
            }
        }

        /// <inheritdoc/>
        public object RTD(string progId, string arg0, params string[] args)
        {
            var arguments = new string[] { progId, string.Empty, arg0 }
                .Concat(args)
                .Select((x, idx) => new FunctionArgument($"p{idx}", x))
                .ToArray();
            var fArgs = new FunctionArguments(arguments);
            IntPtr ptr = IntPtr.Zero;
            using (var pArgs = new StructIntPtr<FunctionArguments>(ref fArgs))
            {
                ptr = AddIn.RtdCall(pArgs.Ptr);
            }
            unsafe
            {
                var result = (XLOPER12*)ptr.ToPointer();
                var obj = result == null ? null : result->ToObject();
                AddIn.AutoFree(ptr);
                return obj;
            }
        }

        /// <inheritdoc/>
        public void SetAsyncResult(IntPtr handle, object result)
        {
            var xlhandle = XLOPER12.FromObject(handle);
            var xlresult = XLOPER12.FromObject(result);
            try
            {
                using (var p1 = new StructIntPtr<XLOPER12>(ref xlhandle))
                using (var p2 = new StructIntPtr<XLOPER12>(ref xlresult))
                    AsyncReturn(p1.Ptr, p2.Ptr);
            }
            finally
            {
                xlresult.Dispose();
                xlhandle.Dispose();
            }
        }

        /// <inheritdoc/>
        public void SetRangeValue(RangeReference range, object value, bool async)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="functions"></param>
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

        private static void AsyncReturn(IntPtr handle, IntPtr result)
        {
            AddIn.AutoFree(AddIn.AsyncReturn(handle, result));
        }
    }
}
