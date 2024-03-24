using System;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public IntPtr Object2IntPtr(object value)
        {
            // TODO
            return IntPtr.Zero;
        }

        public IntPtr Double2IntPtr(double value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr String2IntPtr(string value)
        {
            var len = (ushort)Math.Min(value.Length, XLOPER12.MaxStringLength);
            char* p = (char*)StringValue.ToPointer();
            p[0] = (char) len;
            for (ushort idx = 0; idx  < len; idx++)
                p[idx+1] = value[idx];
            return StringValue;
        }

        public static MethodInfo Result2IntPtr(Type result)
        {
            var flags = BindingFlags.Instance| BindingFlags.Public;
            if (typeof(double) == result)
                return typeof(XlMarshalContext).GetMethod(nameof(Double2IntPtr), flags);
            if (typeof(string) == result)
                return typeof(XlMarshalContext).GetMethod(nameof(String2IntPtr), flags);
            if (typeof(object) == result)
                return typeof(XlMarshalContext).GetMethod(nameof(Object2IntPtr), flags);
            return typeof(XlMarshalContext).GetMethod(nameof(IntPtr2IntPtr), flags);
        }
    }
}
