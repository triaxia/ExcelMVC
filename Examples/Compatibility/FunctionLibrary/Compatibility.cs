using ExcelDna.Integration;
using ExcelMvc.Functions;
using Function.Interfaces;

namespace FunctionLibrary
{
    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public class FunctionAttribute : ExcelDna.Integration.ExcelFunctionAttribute, IFunctionAttribute
    {
        public bool IsAsync { get; set; }
        public new string Category { get => base.Category; set => base.Category = value; }
        public new string Name { get => base.Name; set => base.Name = value; }
        public new string Description { get => base.Description; set => base.Description = value; }
        public new string HelpTopic { get => base.HelpTopic; set => base.HelpTopic = value; }
        public new bool IsVolatile { get => base.IsVolatile; set => base.IsVolatile = value; }
        public new bool IsMacroType { get => base.IsMacroType; set => base.IsMacroType = value; }
        public new bool IsHidden { get => base.IsHidden; set => base.IsHidden = value; }
        public new bool IsThreadSafe { get => base.IsThreadSafe; set => base.IsThreadSafe = value; }
        public new bool IsClusterSafe { get => base.IsClusterSafe; set => base.IsClusterSafe = value; }
        public FunctionAttribute() { }
        public FunctionAttribute(string description) => base.Description = description;
    }

    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    public class ArgumentAttribute : ExcelDna.Integration.ExcelArgumentAttribute, IArgumentAttribute
    {
        public new string Name { get => base.Name; set => base.Name = value; }
        public new string Description { get => base.Description; set => base.Description = value; }
        public ArgumentAttribute() { }
        public ArgumentAttribute(string description) => base.Description = description;
    }

    /// <summary>
    /// Loses this class to lose ExcelDna!
    /// </summary>
    public class ExcelDnaHost : IHost
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
        public bool ExecutingEventRaised { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public Func<Exception, object> ExceptionToFunctionResult { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public Type FunctionAttributeType { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public Type ArgumentAttributeType { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

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
            throw new NotImplementedException();
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
            return ExcelDna.Integration.ExcelDnaUtil.IsInFunctionWizard();
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
            throw new NotImplementedException();
        }

        public void SetAsyncResult(IntPtr handle, object result)
        {
        }

        public void SetRangeValue(RangeReference range, object value, bool async)
        {
        }
    }
}
