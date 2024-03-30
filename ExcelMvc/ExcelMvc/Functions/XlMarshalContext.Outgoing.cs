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

        public IntPtr BoolToIntPtr(bool value)
        {
            *((short*)ShortValue.ToPointer()) = value ? (short) 1 : (short) 0;
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
            *((double*)DoubleValue.ToPointer()) = (double) value;
            return DoubleValue;
        }

        public IntPtr LongToIntPtr(long value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }
        public IntPtr IntToIntPtr(int value)
        {
            *((int*)IntValue.ToPointer()) = value;
            return IntValue;
        }
        
        public IntPtr ShortToIntPtr(short value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr ByteToIntPtr(byte value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
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
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(BoolToIntPtr)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(DoubleToIntPtr)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(DateTimeToIntPtr)) },
                { typeof(decimal), typeof(XlMarshalContext).GetMethod(nameof(DecimalToIntPtr)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(FloatToIntPtr)) },
                { typeof(long), typeof(XlMarshalContext).GetMethod(nameof(LongToIntPtr)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(IntToIntPtr)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(ShortToIntPtr)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(ByteToIntPtr)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(StringToIntPtr)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(ObjectToIntPtr)) },
            };

        public static MethodInfo OutgoingConverter(Type result) =>
            OutgoingConverters.TryGetValue(result, out var value) ? value : OutgoingConverters[(typeof(object))];
    }
}
