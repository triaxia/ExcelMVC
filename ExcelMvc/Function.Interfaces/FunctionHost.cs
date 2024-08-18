using System;
using System.Collections.Generic;
using System.IO;

namespace Function.Interfaces
{
    public class Missing
    {
        private Missing() { }
        public static readonly Missing Value = new Missing();
        public override string ToString() => "";
    }

    public class Empty
    {
        private Empty() { }
        public static readonly Empty Value = new Empty();
        public override string ToString() => "";
    }

    public enum ErrorType : short
    {
        // happen to be Excel error numbers!
        Null = 0,
        Div0 = 7,
        Value = 15,
        Ref = 23,
        Name = 29,
        Num = 36,
        NA = 42,
        Data = 43
    }

    public class NullFunctionHost : IFunctionHost
    {
        /// <inheritdoc/>
        public object Application { get; set; }

        /// <inheritdoc/>
        public object ValueMissing => Missing.Value;

        /// <inheritdoc/>
        public object ValueEmpty => Empty.Value;

        /// <inheritdoc/>
        public object ErrorNull => ErrorType.Null;

        /// <inheritdoc/>
        public object ErrorDiv0 => ErrorType.Div0;

        /// <inheritdoc/>
        public object ErrorValue => ErrorType.Value;

        /// <inheritdoc/>
        public object ErrorRef => ErrorType.Ref;

        /// <inheritdoc/>
        public object ErrorName => ErrorType.Name;

        /// <inheritdoc/>
        public object ErrorNum => ErrorType.Num;

        /// <inheritdoc/>
        public object ErrorNA => ErrorType.NA;

        /// <inheritdoc/>
        public object ErrorData => ErrorType.Data;

        /// <inheritdoc/>
        public int RtdThrottleIntervalMilliseconds { get; set; }

        /// <inheritdoc/>
        public string StatusBarText { get; set; }

        /// <inheritdoc/>
        public bool ExecutingEventRaised { get; set; }

        /// <inheritdoc/>
        public Func<Exception, object> ExceptionToFunctionResult { get; set; }

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
            { ErrorType.Null,"#NULL!" },
            { ErrorType.Div0,"#DIV0!" },
            { ErrorType.Value,"#VALUE!" },
            { ErrorType.Ref,"#REF!" },
            { ErrorType.Name,"#NAME?" },
            { ErrorType.Num,"#NUM!" },
            { ErrorType.NA,"#N/A" },
            { ErrorType.Data,"#Data!" }
        };

        /// <inheritdoc/>
        public string ErrorToString(object value)
        {
            return ErrorMappings.TryGetValue(value, out var mapped) ? mapped : $"{value}";
        }

        /// <inheritdoc/>
        public IntPtr GetAsyncHandle(IntPtr handle)
        {
            return handle;
        }

        /// <inheritdoc/>
        public void SetAsyncValue(IntPtr handle, object value)
        {
        }

        /// <inheritdoc/>
        public RangeReference GetCallerReference()
        {
            return null;
        }

        /// <inheritdoc/>
        public RangeReference GetActiveBookReference(string pageName, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            return null;
        }

        /// <inheritdoc/>
        public RangeReference GetActiveSheetReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            return null;
        }

        /// <inheritdoc/>
        public object GetRangeValue(RangeReference range)
        {
            return null;
        }

        /// <inheritdoc/>
        public void SetRangeValue(RangeReference range, object value, bool async)
        {
        }

        /// <inheritdoc/>
        public RangeReference GetReference(string bookName, string pageName
            , int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            return null;
        }

        /// <inheritdoc/>
        public bool IsInFunctionWizard() => false;

        /// <inheritdoc/>
        public void RaiseExecuting(object sender, ExecutingEventArgs args)
        {
            Executing?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public void RaiseFailed(object sender, ErrorEventArgs args)
        {
            Failed?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public void RaisePosted(object sender, MessageEventArgs args)
        {
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
            return null;
        }

        /// <inheritdoc/>
        public object Rtd(string progId, string server, params string[] args)
        {
            return null;
        }

        /// <inheritdoc/>
        public void RegisterFunctions(FunctionDefinitions functions)
        {
        }

        /// <inheritdoc/>
        public void RaiseRtdUpdated(object sender, RtdServerUpdatedEventArgs args)
        {
            RtdUpdated?.Invoke(sender, args);
        }

        /// <inheritdoc/>
        public void PostAction(Action<object> action, object state)
        {
        }

        /// <inheritdoc/>
        public void PostMacro(Action<object> action, object state)
        {
        }

        public object Run(int function, params object[] args)
        {
            return null;
        }

        /// <inheritdoc/>
        public string Version => "";

        /// <inheritdoc/>
        public bool IsIdeOpen => false;

        /// <inheritdoc/>
        public string ModuleFileName => "";
    }

    /// <summary>
    /// </summary>
    public static class FunctionHost
    {
        /// <summary>
        /// Gets/Sets the implementation of <see cref="IFunctionHost"/>.
        /// </summary>
        public static IFunctionHost Instance { get; set; } = new NullFunctionHost();
    }
}
