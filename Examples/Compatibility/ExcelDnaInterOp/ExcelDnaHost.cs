using ExcelDna.Integration;
using ExcelMvc.Functions;
using Function.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelDnaInterOp
{
    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    public class ExcelDnaHost : IFunctionHost
    {
        public object Underlying => ExcelDnaUtil.Application;

        public object ValueMissing => ExcelDna.Integration.ExcelMissing.Value;

        public object ValueEmpty => ExcelDna.Integration.ExcelEmpty.Value;

        public object ErrorNull => ExcelDna.Integration.ExcelError.ExcelErrorNull;

        public object ErrorDiv0 => ExcelDna.Integration.ExcelError.ExcelErrorDiv0;

        public object ErrorValue => ExcelDna.Integration.ExcelError.ExcelErrorValue;

        public object ErrorRef => ExcelDna.Integration.ExcelError.ExcelErrorRef;

        public object ErrorName => ExcelDna.Integration.ExcelError.ExcelErrorName;

        public object ErrorNA => ExcelDna.Integration.ExcelError.ExcelErrorNA;

        public object ErrorData => ExcelDna.Integration.ExcelError.ExcelErrorGettingData;

        public int RtdThrottleIntervalMilliseconds { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public string StatusBarText { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool ExecutingEventRaised { get; set; } = false;

        public Func<Exception, object> ExceptionToFunctionResult { get; set; }
        public Type FunctionAttributeType { get; set; } = typeof(FunctionAttribute);
        public Type ArgumentAttributeType { get; set; } = typeof(ArgumentAttribute);

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
            throw new NotImplementedException();
        }

        public RangeReference GetActivePageReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            throw new NotImplementedException();
        }

        public IntPtr GetAsyncHandle(IntPtr handle)
        {
            throw new NotImplementedException();
        }

        public RangeReference GetCallerReference()
        {
            /*
            var reference = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            return new RangeReference(reference.)
            */
            return null;
        }

        public object GetRangeValue(RangeReference range)
        {
            throw new NotImplementedException();
        }

        public RangeReference GetReference(string bookName, string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            throw new NotImplementedException();
        }

        public bool IsInFunctionWizard()
        {
            return ExcelDnaUtil.IsInFunctionWizard();
        }

        public void RaiseExecuting(object sender, ExecutingEventArgs args)
        {
        }

        public void RaiseFailed(object sender, ErrorEventArgs args)
        {
        }

        public void RaisePosted(object sender, MessageEventArgs args)
        {
        }

        public void RaiseRegistering(object sender, RegisteringEventArgs args)
        {
        }

        public void RaiseRtdUpdated(object sender, RtdServerUpdatedEventArgs args)
        {
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
        }

        public void SetRangeValue(RangeReference range, object value, bool async)
        {
        }
    }
}
