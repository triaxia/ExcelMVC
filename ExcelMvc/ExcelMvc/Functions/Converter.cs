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
                GetDefaultValue(info) : Marshal.PtrToStructure<XLOPER12_num>(incoming).num;
            return result;
        }

        public static void ConvertOutging(object outgoing, ref IntPtr result)
        {
            XLOPER12_num r;
            r.xltype = 1;
            r.num = (double)outgoing;
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