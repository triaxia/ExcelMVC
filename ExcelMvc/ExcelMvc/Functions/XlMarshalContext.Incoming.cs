using System;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public static IntPtr IntPtr2IntPtr(IntPtr value) => value;
        public static double IntPtr2Double(IntPtr value) => *(double*)value.ToPointer();
        public static string IntPtr2String(IntPtr value)
        {
            if (value == IntPtr.Zero) return null;

            char* p = (char*)value.ToPointer();
            var len = (short)p[0];
            return len == 0 ? string.Empty : new string(p, 1, len);
        }

        public static MethodInfo IntPtr2Parameter(Type parameter)
        {
            var flags = BindingFlags.Static | BindingFlags.Public;
            if (typeof(double) == parameter)
                return typeof(XlMarshalContext).GetMethod(nameof(IntPtr2Double), flags);
            if (typeof(string) == parameter)
                return typeof(XlMarshalContext).GetMethod(nameof(IntPtr2String), flags);

            return typeof(XlMarshalContext).GetMethod(nameof(IntPtr2IntPtr), flags);
        }
    }
}
