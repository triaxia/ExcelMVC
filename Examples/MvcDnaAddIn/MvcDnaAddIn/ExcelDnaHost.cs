using ExcelDna.Integration;
using ExcelMvc.Functions;
using Function.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelAddIn
{
    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    public class ExcelDnaHost : IFunctionHost
    {
        private IFunctionHost DelegateHost { get; }

        public ExcelDnaHost()
        {
            DelegateHost = new ExcelFunctionHost
            {
                Underlying = ExcelDnaUtil.Application
            };
        }

        public object Underlying { get; set; } = ExcelDnaUtil.Application;

        public object ValueMissing => ExcelDna.Integration.ExcelMissing.Value;

        public object ValueEmpty => ExcelDna.Integration.ExcelEmpty.Value;

        public object ErrorNull => ExcelDna.Integration.ExcelError.ExcelErrorNull;

        public object ErrorDiv0 => ExcelDna.Integration.ExcelError.ExcelErrorDiv0;

        public object ErrorValue => ExcelDna.Integration.ExcelError.ExcelErrorValue;

        public object ErrorRef => ExcelDna.Integration.ExcelError.ExcelErrorRef;

        public object ErrorName => ExcelDna.Integration.ExcelError.ExcelErrorName;

        public object ErrorNA => ExcelDna.Integration.ExcelError.ExcelErrorNA;

        public object ErrorData => ExcelDna.Integration.ExcelError.ExcelErrorGettingData;

        public int RtdThrottleIntervalMilliseconds
        {
            get => DelegateHost.RtdThrottleIntervalMilliseconds;
            set => DelegateHost.RtdThrottleIntervalMilliseconds = value;
        }
        public string StatusBarText
        {
            get => DelegateHost.StatusBarText;
            set => DelegateHost.StatusBarText = value;
        }
        public bool ExecutingEventRaised { get; set; } = false;

        public Func<Exception, object> ExceptionToFunctionResult { get; set; }
            = e => $"{e.Message}";

        public Type FunctionAttributeType { get; set; } = typeof(FunctionAttribute);
        public Type ArgumentAttributeType { get; set; } = typeof(ArgumentAttribute);

        public string Version => DelegateHost.Version;
        public bool IsIdeOpen => DelegateHost.IsIdeOpen;

        public event EventHandler<RtdServerUpdatedEventArgs> RtdUpdated;
        public event EventHandler<MessageEventArgs> Posted;
        public event EventHandler<RegisteringEventArgs> Registering;
        public event EventHandler<ErrorEventArgs> Failed;
        public event EventHandler<ExecutingEventArgs> Executing;

        public static Dictionary<object, string> Mappings = new Dictionary<object, string>
        {
            { ExcelDna.Integration.ExcelError.ExcelErrorNull,"#NULL!" },
            { ExcelDna.Integration.ExcelError.ExcelErrorDiv0,"#DIV0!" },
            { ExcelDna.Integration.ExcelError.ExcelErrorValue,"#VALUE!" },
            { ExcelDna.Integration.ExcelError.ExcelErrorRef,"#REF!" },
            { ExcelDna.Integration.ExcelError.ExcelErrorName,"#NAME?" },
            { ExcelDna.Integration.ExcelError.ExcelErrorNum,"#NUM!" },
            { ExcelDna.Integration.ExcelError.ExcelErrorNA,"#N/A" },
            { ExcelDna.Integration.ExcelError.ExcelErrorGettingData,"#Data!" },
            { ExcelDna.Integration.ExcelMissing.Value, $"{ExcelDna.Integration.ExcelMissing.Value}" },
            { ExcelDna.Integration.ExcelEmpty.Value, $"{ExcelDna.Integration.ExcelEmpty.Value}" }
        };

        public string ErrorToString(object value)
        {
            return Mappings.TryGetValue(value, out var mapped) ? mapped : $"{value}";
        }

        public RangeReference GetActiveBookReference(string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            return DelegateHost.GetActiveBookReference(pageName, rowFirst, rowLast, columnFirst, columnLast);
        }

        public RangeReference GetActivePageReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            return DelegateHost.GetActivePageReference(rowFirst, rowLast, columnFirst, columnLast);
        }

        public IntPtr GetAsyncHandle(IntPtr handle)
        {
            return DelegateHost.GetAsyncHandle(handle);
        }

        public RangeReference GetCallerReference()
        {
            return DelegateHost.GetCallerReference();
        }

        public object GetRangeValue(RangeReference range)
        {
            return DelegateHost.GetRangeValue(range);
        }

        public RangeReference GetReference(string bookName, string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            return DelegateHost.GetReference(bookName, pageName, rowFirst, rowLast, columnFirst, columnLast);
        }

        public bool IsInFunctionWizard()
        {
            return ExcelDnaUtil.IsInFunctionWizard();
        }

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
            DelegateHost.SetAsyncResult(handle, result);
        }

        public void SetRangeValue(RangeReference range, object value, bool async)
        {
            DelegateHost.SetRangeValue(range, value, async);
        }

        public void Post(Action<object> action, object state)
        {
            ExcelAsyncUtil.QueueAsMacro(x => action(x), state);
        }
    }
}
