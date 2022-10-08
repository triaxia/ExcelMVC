using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class Converter
    {
        public static object ConvertIncoming(IntPtr incoming, ParameterInfo info)
        {
            var result = incoming == IntPtr.Zero ?
                GetDefaultValue(info) : Marshal.PtrToStructure<XLOPER12>(incoming).num;
            return result;
        }

        public static void ConvertOutging(object outgoing, MethodInfo method, ref IntPtr result)
        {
            var r = new XLOPER12((double)outgoing);
            Marshal.StructureToPtr(r, result, false);
        }

        private static object GetDefaultValue(ParameterInfo info)
        {
            var type = info.ParameterType;
            return type.IsValueType ? Activator.CreateInstance(type) : null;
        }
    }
}
;