using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public static IntPtr IntPtrToIntPtr(IntPtr value) => value;
        public static bool IntPtrToBoolean(IntPtr value) => *(short*)value.ToPointer() != 0;
        public static double IntPtrToDouble(IntPtr value) => *(double*)value.ToPointer();
        public static DateTime IntPtrToDateTime(IntPtr value) => DateTime.FromOADate(*(double*)value.ToPointer());
        public static float IntPtrToSingle(IntPtr value) => (float)*(double*)value.ToPointer();
        public static int IntPtrToInt32(IntPtr value) => *(int*)value.ToPointer();
        public static uint IntPtrToUInt32(IntPtr value) => (uint)*(int*)value.ToPointer();
        public static short IntPtrToInt16(IntPtr value) => *(short*)value.ToPointer();
        public static ushort IntPtrToUInt16(IntPtr value) => (ushort) *(short*)value.ToPointer();
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

        public static DateTime[] IntPtrToDateTimeArray(IntPtr value)
        {
            if (value == IntPtr.Zero)
                return null;

            var cells = IntPtrToDoubleArray(value);
            if (cells.Length == 0)
                return new DateTime[] { };

            var result = new DateTime[cells.Length];
            for (var i = 0; i < cells.Length; i++)
                result[i] = DateTime.FromOADate(cells[i]);
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

        public static DateTime[,] IntPtrToDateTimeMatrix(IntPtr value)
        {
            if (value == IntPtr.Zero)
                return null;
            var cells = IntPtrToDoubleMatrix(value);

            var rows = cells.GetLength(0);
            var cols = cells.GetLength(1);
            if (rows == 0 || cols == 0)
                return new DateTime[,] { };
            var result = new DateTime[rows, cols];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    result[row, col] = DateTime.FromOADate(cells[row, col]);
            return result;
        }

        public static object IntPtrToObject(IntPtr value)
        {
            XLOPER12* p = (XLOPER12*)value.ToPointer();
            return p->ToObject();
        }

        private static readonly Dictionary<Type, MethodInfo> IncomingConverters
            = new Dictionary<Type, MethodInfo>()
            {
                { typeof(IntPtr), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToIntPtr)) },
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToBoolean)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDouble)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTime)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToSingle)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt32)) },
                { typeof(uint), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUInt32)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt16)) },
                { typeof(ushort), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUInt16)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToByte)) },
                { typeof(sbyte), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToSByte)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToString)) },
                { typeof(double[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDoubleArray)) },
                { typeof(double[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDoubleMatrix)) },
                { typeof(DateTime[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTimeArray)) },
                { typeof(DateTime[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTimeMatrix)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToObject)) }
            };

        public static MethodInfo IncomingConverter(Type result) =>
            IncomingConverters.TryGetValue(result, out var value) ? value : IncomingConverters[(typeof(object))];
    }
}
