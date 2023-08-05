using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Explicit)]
    unsafe struct XlString12
    {
        [FieldOffset(0)]
        public ushort Length;
        [FieldOffset(2)]
        public char *Data;
    }

    [StructLayout(LayoutKind.Explicit)]
    unsafe public struct XLOPER12
    {
        [FieldOffset(0)] public double num;
        [FieldOffset(0)] public int w;
        [FieldOffset(0)] public int xbool;
        [FieldOffset(0)] private void *str;
        [FieldOffset(24)] public uint xltype;

        public XLOPER12(double v)
        {
            str = null;
            w = 0;
            xbool = 0;
            num = v;
            xltype = (uint)XlTypes.xltypeNum;
        }

        public XLOPER12(int v)
        {
            str = null;
            num = 0;
            xbool = 0;
            xltype = (uint)XlTypes.xltypeInt;
            w = v;
        }

        public XLOPER12(bool v)
        {
            str = null;
            num = 0;
            w = 0;
            xltype = (uint)XlTypes.xltypeBool;
            xbool = v ? -1 : 0;
        }

        public XLOPER12(string v)
        {
            num = 0;
            w = 0;
            xbool =0;
            xltype = (uint)XlTypes.xltypeStr;

            var x = new XlString12
            {
                Length = (ushort)v.Length,
                Data = (char*)Marshal.AllocCoTaskMem(v.Length)
            };

            var idx = 0;
            foreach (char c in v) x.Data[idx++] = c;

            using (var result = new StructIntPtr<XlString12>(ref x))
                str = (void*) result.Detach();
        }

        public string Str()
        {
            if (xltype != (uint)XlTypes.xltypeStr)
                return null;

            var value = Marshal.PtrToStringUni(new IntPtr(str));
            return value.Substring(1, value[0]);
        }
    }
}
