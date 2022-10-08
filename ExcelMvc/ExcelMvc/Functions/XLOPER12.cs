using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Explicit)]
    public struct XlString
    {
        [FieldOffset(0)]
        public ushort Length;
        [FieldOffset(2)]
        public IntPtr Data;
    }
   
    [StructLayout(LayoutKind.Explicit)]
    public struct XLOPER12
    {
        [FieldOffset(0)] public double num;
        [FieldOffset(0)] public int w;
        [FieldOffset(0)] public int xbool;
        //[FieldOffset(0)] public XlString str;
        [FieldOffset(24)] public uint xltype;

        public static void Make(double v, out XLOPER12 op)
        {
            //op.str.Length = 0;
            //op.str.Data = IntPtr.Zero;
            op.w = 0;
            op.xbool = 0;
            op.num = v;
            op.xltype = (uint)XlTypes.xltypeNum;
        }

        //public static void Make(int v, out XLOPER12 op)
        //{
        //    op.str = null;
        //    op.num = 0;
        //    op.xbool = 0;
        //    op.xltype = (uint)XlTypes.xltypeInt;
        //    op.w = v;
        //}

        //public static void Make(bool v, out XLOPER12 op)
        //{
        //    op.str = null;
        //    op.num = 0;
        //    op.w = 0;
        //    op.xltype = (uint)XlTypes.xltypeBool;
        //    op.xbool = v ? -1 : 0;
        //}
    }
}
