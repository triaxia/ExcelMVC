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
        public static ulong IntPtrToULong(IntPtr value) => (ulong)*(double*)value.ToPointer();
        public static int IntPtrToInt(IntPtr value) => *(int*)value.ToPointer();
        public static uint IntPtrToUInt(IntPtr value) => (uint)*(double*)value.ToPointer();
        public static short IntPtrToShort(IntPtr value) => *(short*)value.ToPointer();
        public static ushort IntPtrToUShort(IntPtr value) => (ushort) * (int*)value.ToPointer();
        public static byte IntPtrToByte(IntPtr value) => (byte)*(short*)value.ToPointer();
        public static sbyte IntPtrToSByte(IntPtr value) => (sbyte)*(short*)value.ToPointer();
        public static string IntPtrToString(IntPtr value)
        {
            if (value == IntPtr.Zero)
                return null;

            char* p = (char*)value.ToPointer();
            return new string(p);
        }

        public static double[] IntPtrToDoubleArray(IntPtr value)
        {
            if (value == IntPtr.Zero)
                return null;

            int* p = (int*)value.ToPointer();
            var rows = p[0];
            var cols = p[1];
            if (rows == 0 || cols == 0)
                return new double[] { };
            var len = rows * cols;
            var result = new double[len];
            double *x = (double*) &p[2];
            for (var i = 0; i < len; i++)
                result[i] = x[i];
            return result;
        }

        public static double[,] IntPtrToDoubleMatrix(IntPtr value)
        {
            if (value == IntPtr.Zero)
                return null;

            int* p = (int*)value.ToPointer();
            var rows = p[0];
            var cols = p[1];
            if (rows == 0 || cols == 0)
                return new double[,] { };
            var len = rows * cols;
            var result = new double[rows, cols];
            double* x = (double*)&p[2];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    result[row, col] = x[row * cols + col];
            return result;
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
                { typeof(ulong), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToULong)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt)) },
                { typeof(uint), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUInt)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToShort)) },
                { typeof(ushort), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUShort)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToByte)) },
                { typeof(sbyte), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToSByte)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToString)) },
                { typeof(double[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDoubleArray)) },
                { typeof(double[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDoubleMatrix)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToObject)) }
            };

        public static MethodInfo IncomingConverter(Type result) =>
            IncomingConverters.TryGetValue(result, out var value) ? value : IncomingConverters[(typeof(object))];
    }
}
