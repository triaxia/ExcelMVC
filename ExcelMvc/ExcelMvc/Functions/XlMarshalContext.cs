using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelMvc.Functions
{
    public unsafe class XlMarshalContext
    {
        private readonly IntPtr DoubleValue;

        // thread affinity for return pointers...
        private readonly static ThreadLocal<XlMarshalContext> MarshalContext
            = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetMarshalContext() => MarshalContext.Value;

        public XlMarshalContext()
        {
            var size = Marshal.SizeOf(typeof(double));
            DoubleValue = Marshal.AllocCoTaskMem(size);
        }

        ~XlMarshalContext()
        {
            Marshal.FreeCoTaskMem(DoubleValue);
        }

        public static double IntPtr2Double(IntPtr value)
        {
            return *(double*)value.ToPointer();
        }

        public static MethodInfo IntPtr2DoubleMethod =>
            typeof(XlMarshalContext).GetMethod(nameof(IntPtr2Double));

        public static string IntPtr2String(IntPtr value)
        {
            char* p = (char*)value.ToPointer();
            var len = (short)p[0];
            return len == 0 ? string.Empty : new string(p, 1, len);
        }

        public static MethodInfo IntPtr2StringMethod =>
            typeof(XlMarshalContext).GetMethod(nameof(IntPtr2String));

        public static MethodInfo IntPtr2ParameterMethod(Type parameter)
        {
            if (typeof(double) == parameter)
                return IntPtr2DoubleMethod;
            if (typeof(string) == parameter)
                return IntPtr2StringMethod;
            return IntPtr2StringMethod;
        }
    }
}
