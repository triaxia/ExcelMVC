using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public IntPtr ObjectToIntPtr(object value)
        {
            // TODO
            return IntPtr.Zero;
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
            *((decimal*)DecimalValue.ToPointer()) = value;
            return DecimalValue;
        }

        public IntPtr LongToIntPtr(long value)
        {
            *((long*)LongValue.ToPointer()) = value;
            return LongValue;
        }

        public IntPtr ULongToIntPtr(ulong value)
        {
            *((ulong*)LongValue.ToPointer()) = value;
            return LongValue;
        }

        public IntPtr IntToIntPtr(int value)
        {
            *((long*)LongValue.ToPointer()) = value;
            return LongValue;
        }

        public IntPtr UIntToIntPtr(uint value)
        {
            *((long*)LongValue.ToPointer()) = value;
            return LongValue;
        }

        public IntPtr ShortToIntPtr(short value)
        {
            *((long*)LongValue.ToPointer()) = value;
            return LongValue;
        }

        public IntPtr UShortToIntPtr(ushort value)
        {
            *((long*)LongValue.ToPointer()) = value;
            return LongValue;
        }

        public IntPtr ByteToIntPtr(byte value)
        {
            *((long*)LongValue.ToPointer()) = value;
            return LongValue;
        }

        public IntPtr StringToIntPtr(string value)
        {
            var len = (ushort)Math.Min(value.Length, XLOPER12.MaxStringLength);
            char* p = (char*)StringValue.ToPointer();
            p[0] = (char)len;
            for (ushort idx = 0; idx < len; idx++)
                p[idx + 1] = value[idx];
            return StringValue;
        }

        private static readonly Dictionary<Type, MethodInfo> OutgoingConverters
            = new Dictionary<Type, MethodInfo>()
            {
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
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(StringToIntPtr)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(ObjectToIntPtr)) },
            };

        public static MethodInfo OutgoingConverter(Type result) =>
            OutgoingConverters.TryGetValue(result, out var value) ? value : OutgoingConverters[(typeof(object))];
    }
}
