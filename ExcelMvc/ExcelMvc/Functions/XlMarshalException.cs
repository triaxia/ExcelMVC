using System;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public class XlMarshalExceptionEventArgs : EventArgs
    {
        public Exception Exception { get; }
        public object Value { get; set; }
        public XlMarshalExceptionEventArgs(Exception ex)
        {
            Exception = ex;
            Value = XlErrors.xlerrValue;
        }
    }

    public static class XlMarshalException
    {
        public static event EventHandler<EventArgs> Failed;
        public static object HandleUnhandledException(object ex)
        {
            if (Failed == null) return XlErrors.xlerrValue;
            try
            {
                var args = new XlMarshalExceptionEventArgs((Exception)ex);
                Failed.Invoke(null, args);
                return args.Value;
            }
            catch
            {
                return XlErrors.xlerrValue;
            }
        }

        public static MethodInfo HandlerMethod =>
            typeof(XlMarshalException).GetMethod(nameof(XlMarshalException.HandleUnhandledException)
                , BindingFlags.Static | BindingFlags.Public);
    }
}
