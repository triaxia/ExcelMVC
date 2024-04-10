using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public IntPtr ObjectToIntPtr(object value)
        {
            XLOPER12* p = (XLOPER12*)ObjectValue.ToPointer();
            p->Init(value);
            return ObjectValue;
        }

        public IntPtr BooleanToIntPtr(bool value)
        {
            *((short*)ShortValue.ToPointer()) = value ? (short)1 : (short)0;
            return ShortValue;
        }

        public IntPtr DoubleToIntPtr(double value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr DateTimeToIntPtr(DateTime value)
        {
            *((double*)DoubleValue.ToPointer()) = value.ToOADate();
            return DoubleValue;
        }

        public IntPtr SingleToIntPtr(float value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr Int32ToIntPtr(int value)
        {
            *((int*)IntValue.ToPointer()) = value;
            return IntValue;
        }

        public IntPtr UInt32ToIntPtr(uint value)
        {
            *((int*)IntValue.ToPointer()) =(int) value;
            return IntValue;
        }

        public IntPtr Int16ToIntPtr(short value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr UInt16ToIntPtr(ushort value)
        {
            *((short*)ShortValue.ToPointer()) = (short) value;
            return ShortValue;
        }

        public IntPtr ByteToIntPtr(byte value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr SByteToIntPtr(sbyte value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr StringToIntPtr(string value)
        {
            Marshal.FreeCoTaskMem(StringValue);
            var len = value?.Length ?? 0;
            StringValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(char)) * (len + 1));
            char* p = (char*)StringValue.ToPointer();
            p[len] = '\0';
            for (var idx = 0; idx < len; idx++)
                p[idx] = value[idx];
            return StringValue;
        }

        public IntPtr DoubleArrayToIntPtr(double[] value)
        {
            Marshal.FreeCoTaskMem(DoubleArrayValue);
            var len = value?.Length ?? 0;
            DoubleArrayValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)) * 2 +
                Marshal.SizeOf(typeof(double)) * len);
            int* p = (int*)DoubleArrayValue.ToPointer();
            p[0] = 1;
            p[1] = len;

            double* d = (double*)&p[2];
            for (var idx = 0; idx < len; idx++)
                d[idx] = value[idx];
            return DoubleArrayValue;
        }

        public IntPtr DateTimeArrayToIntPtr(DateTime[] value)
        {
            return DoubleArrayToIntPtr(value.Select(x=>x.ToOADate()).ToArray());
        }

        public IntPtr DoubleMatrixToIntPtr(double[,] value)
        {
            Marshal.FreeCoTaskMem(DoubleArrayValue);
            var rows = value?.GetLength(0) ?? 0;
            var cols = value?.GetLength(1) ?? 0;
            DoubleArrayValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)) * 2 +
                Marshal.SizeOf(typeof(double)) * rows * cols);
            int* p = (int*)DoubleArrayValue.ToPointer();
            p[0] = rows;
            p[1] = cols;

            double* d = (double*)&p[2];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    d[row * cols + col] = value[row, col];
            return DoubleArrayValue;
        }

        public IntPtr DateTimeMatrixToIntPtr(DateTime[,] value)
        {
            var rows = value?.GetLength(0) ?? 0;
            var cols = value?.GetLength(1) ?? 0;
            if (rows == 0 || cols == 0)
                return DoubleMatrixToIntPtr(new double[,] { });
            var cells = new double[rows, cols];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    cells[row, col] = value[row, col].ToOADate();
            return DoubleMatrixToIntPtr(cells);
        }

        private static readonly Dictionary<Type, MethodInfo> OutgoingConverters
            = new Dictionary<Type, MethodInfo>()
            {
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(BooleanToIntPtr)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(DoubleToIntPtr)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(DateTimeToIntPtr)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(SingleToIntPtr)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(Int32ToIntPtr)) },
                { typeof(uint), typeof(XlMarshalContext).GetMethod(nameof(UInt32ToIntPtr)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(Int16ToIntPtr)) },
                { typeof(ushort), typeof(XlMarshalContext).GetMethod(nameof(UInt16ToIntPtr)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(ByteToIntPtr)) },
                { typeof(sbyte), typeof(XlMarshalContext).GetMethod(nameof(SByteToIntPtr)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(StringToIntPtr)) },
                { typeof(double[]), typeof(XlMarshalContext).GetMethod(nameof(DoubleArrayToIntPtr)) },
                { typeof(double[,]), typeof(XlMarshalContext).GetMethod(nameof(DoubleMatrixToIntPtr)) },
                { typeof(DateTime[]), typeof(XlMarshalContext).GetMethod(nameof(DateTimeArrayToIntPtr)) },
                { typeof(DateTime[,]), typeof(XlMarshalContext).GetMethod(nameof(DateTimeMatrixToIntPtr)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(ObjectToIntPtr)) },
            };

        public static MethodInfo OutgoingConverter(Type result) =>
            OutgoingConverters.TryGetValue(result, out var value) ? value : OutgoingConverters[(typeof(object))];
    }
}
