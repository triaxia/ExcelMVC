using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class Converter
    {
        public static object ConvertIncoming(IntPtr incoming, ParameterInfo info)
        {
            if (incoming == IntPtr.Zero) return GetDefaultValue(info);
            var x = Marshal.PtrToStructure<XLOPER12>(incoming);
            return x.xx();
            //return x.num;
        }

        public static void ConvertOutging(object outgoing, MethodInfo method, ref IntPtr result)
        {
            var r = new XLOPER12(outgoing.ToString());
            Marshal.StructureToPtr(r, result, false);
        }

        private static object GetDefaultValue(ParameterInfo info)
        {
            var type = info.ParameterType;
            return type.IsValueType ? Activator.CreateInstance(type) : null;
        }
    }
}