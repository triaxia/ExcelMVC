using System;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class Converter
    {
        public static object FromIncoming(XLOPER12? value, ParameterInfo info)
        {
            if (value == null)
                return GetDefaultValue(info);
            var result = XLOPER12.ToObject(value.Value);
            if (result == null)
                return GetDefaultValue(info);
            if (info.ParameterType == typeof(object))
                return result;
            if (info.ParameterType == result.GetType())
                return result;
            // TODO
            return Convert.ChangeType(result, info.ParameterType);
        }

        public static void ToOutgoing(object outgoing, ref IntPtr result, MethodInfo method)
        {
            XLOPER12.ToIntPtr(XLOPER12.FromObject(outgoing), ref result);
        }

        private static object GetDefaultValue(ParameterInfo info)
        {
            var type = info.ParameterType;
            return type.IsValueType ? Activator.CreateInstance(type) : null;
        }
    }
}