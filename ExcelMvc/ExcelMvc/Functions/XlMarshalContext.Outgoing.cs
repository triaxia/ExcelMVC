using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public IntPtr ObjectToIntPtr(object value)
        {
            // TODO
            return IntPtr.Zero;
        }

        public IntPtr BoolToIntPtr(bool value)
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

        public IntPtr FloatToIntPtr(float value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr DecimalToIntPtr(decimal value)
        {
            *((double*)DoubleValue.ToPointer()) = (double)value;
            return DoubleValue;
        }

        public IntPtr LongToIntPtr(long value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr ULongToIntPtr(ulong value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr IntToIntPtr(int value)
        {
            *((int*)IntValue.ToPointer()) = value;
            return IntValue;
        }

        public IntPtr UIntToIntPtr(uint value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr ShortToIntPtr(short value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr UShortToIntPtr(ushort value)
        {
            *((int*)IntValue.ToPointer()) = value;
            return IntValue;
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

        private static readonly Dictionary<Type, MethodInfo> OutgoingConverters
            = new Dictionary<Type, MethodInfo>()
            {
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(BoolToIntPtr)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(DoubleToIntPtr)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(DateTimeToIntPtr)) },
                { typeof(decimal), typeof(XlMarshalContext).GetMethod(nameof(DecimalToIntPtr)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(FloatToIntPtr)) },
                { typeof(long), typeof(XlMarshalContext).GetMethod(nameof(LongToIntPtr)) },
                { typeof(ulong), typeof(XlMarshalContext).GetMethod(nameof(ULongToIntPtr)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(IntToIntPtr)) },
                { typeof(uint), typeof(XlMarshalContext).GetMethod(nameof(UIntToIntPtr)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(ShortToIntPtr)) },
                { typeof(ushort), typeof(XlMarshalContext).GetMethod(nameof(UShortToIntPtr)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(ByteToIntPtr)) },
                { typeof(sbyte), typeof(XlMarshalContext).GetMethod(nameof(SByteToIntPtr)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(StringToIntPtr)) },
                { typeof(double[]), typeof(XlMarshalContext).GetMethod(nameof(DoubleArrayToIntPtr)) },
                { typeof(double[,]), typeof(XlMarshalContext).GetMethod(nameof(DoubleMatrixToIntPtr)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(ObjectToIntPtr)) },
            };

        public static MethodInfo OutgoingConverter(Type result) =>
            OutgoingConverters.TryGetValue(result, out var value) ? value : OutgoingConverters[(typeof(object))];
    }
}
