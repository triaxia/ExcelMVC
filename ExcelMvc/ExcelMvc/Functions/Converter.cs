using System;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class Converter
    {
        public static object FromIncoming(XLOPER12? value, ParameterInfo info)
        {
            if (value == null) return GetDefaultValue(info);
            var result = XLOPER12.ToObject(value.Value);
            return info.ParameterType == typeof(object) ? result : Convert.ChangeType(result, info.ParameterType);
        }

        public static void ToOutgoing(object outgoing, ref IntPtr result, MethodInfo method)
        {
            XLOPER12.ToIntPtr(XLOPER12.FromObject(outgoing),ref result);
        }

        private static object GetDefaultValue(ParameterInfo info)
        {
            var type = info.ParameterType;
            return type.IsValueType ? Activator.CreateInstance(type) : null;
        }
    }
}