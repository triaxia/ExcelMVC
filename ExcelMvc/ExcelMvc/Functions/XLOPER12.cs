using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Explicit)]
    unsafe public struct XLOPER12
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

        unsafe public XLOPER12(string v)
        {
            num = 0;
            w = 0;
            xbool = 0;
            xltype = (uint)XlTypes.xltypeStr;
            char x = (char) v.Length;
            str = (void*)Marshal.StringToBSTR(v);
        }
    }
}
