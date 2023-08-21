using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Explicit)]
    unsafe public struct XLOPER12 : IDisposable
    {
        [FieldOffset(0)] public double num;
        [FieldOffset(0)] public int w;
        [FieldOffset(0)] public int xbool;
        [FieldOffset(0)] public void* str;
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
            xbool = 0;
            xltype = (uint)XlTypes.xltypeStr;
            str = (void*)Marshal.AllocCoTaskMem((v.Length + 1) * sizeof(char));
            char* p = (char*)str;
            p[0] =(char) v.Length;
            for (var idx = 1; idx <= v.Length; idx++)
                p[idx] = v[idx - 1];
        }

        public string xx()
        {
            char* p =(char *) str;
            var length = p[0];
            if (length == 0)
                return String.Empty;
            var d = new char[length];
            for (var idx = 1; idx <= length; idx++)
                d[idx - 1] = p[idx];
            return new string(d);
        }

        public void Dispose()
        {
            if (str != null)
                Marshal.FreeCoTaskMem((IntPtr)str);
        }
    }
}
