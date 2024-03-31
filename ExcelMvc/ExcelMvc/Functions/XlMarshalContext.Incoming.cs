using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public static IntPtr IntPtrToIntPtr(IntPtr value) => value;
        public static bool IntPtrToBool(IntPtr value) => *(short*)value.ToPointer() != 0;
        public static double IntPtrToDouble(IntPtr value) => *(double*)value.ToPointer();
        public static DateTime IntPtrToDateTime(IntPtr value) => DateTime.FromOADate(*(double*)value.ToPointer());
        public static float IntPtrToFloat(IntPtr value) => (float)*(double*)value.ToPointer();
        public static decimal IntPtrToDecimal(IntPtr value) => (decimal)*(double*)value.ToPointer();
        public static long IntPtrToLong(IntPtr value) => (long)*(double*)value.ToPointer();
        public static int IntPtrToInt(IntPtr value) => *(int*)value.ToPointer();
        public static short IntPtrToShort(IntPtr value) => *(short*)value.ToPointer();
        public static byte IntPtrToByte(IntPtr value) => (byte)*(short*)value.ToPointer();
        public static string IntPtrToString(IntPtr value)
        {
            if (value == IntPtr.Zero) return null;

            short* p = (short*)value.ToPointer();
            var len =0;
            while (p[len] != 0) len++;
            return len == 0 ? string.Empty : new string((char*)p, 0, len);
        }

        public static object IntPtrToObject(IntPtr value)
        {
            return null;
        }

        private static readonly Dictionary<Type, MethodInfo> IncomingConverters
            = new Dictionary<Type, MethodInfo>()
            {
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToBool)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDouble)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTime)) },
                { typeof(decimal), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDecimal)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToFloat)) },
                { typeof(long), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToLong)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToShort)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToByte)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToString)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToObject)) },
            };

        public static MethodInfo IncomingConverter(Type result) =>
            IncomingConverters.TryGetValue(result, out var value) ? value : IncomingConverters[(typeof(object))];
    }
}
