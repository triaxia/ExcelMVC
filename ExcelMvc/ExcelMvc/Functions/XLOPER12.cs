using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Sequential)]
    unsafe public struct XLArray
    {
        public XLOPER12* lparray;
        public int rows;
        public int columns;
    }

    [StructLayout(LayoutKind.Explicit)]
    unsafe public struct XLOPER12 : IDisposable
    {
        [FieldOffset(0)] public double num;
        [FieldOffset(0)] public int w;
        [FieldOffset(0)] public int err;
        [FieldOffset(0)] public void* any;
        [FieldOffset(0)] XLArray array;
        [FieldOffset(24)] public uint xltype;

        public static XLOPER12 FromObject(object value)
        {
            return new XLOPER12(value);
        }

        public static XLOPER12? FromIntPtr(IntPtr value)
        {
            return value ==IntPtr.Zero ? null : (XLOPER12 ?) Marshal.PtrToStructure<XLOPER12>(value);
        }

        public static object ToObject(XLOPER12 value)
        {
            return value.ToObject();
        }

        public static void ToIntPtr(XLOPER12 value, ref IntPtr result)
        {
            Marshal.StructureToPtr(value, result, false);
        }

        public void Dispose()
        {
            if (any != null)
                Marshal.FreeHGlobal((IntPtr)any);
        }

        private XLOPER12(object value)
        {
            num = 0;
            w = 0;
            any = null;
            array = new XLArray
            {
                rows = 0,
                columns = 0,
                lparray = null
            };
            xltype = (uint)XlTypes.xltypeNil;
            err = 0;
            Init(value);
        }

        private void Init(object value)
        {
            num = 0;
            w = 0;
            any = null;
            array = new XLArray
            {
                rows = 0,
                columns = 0,
                lparray = null
            };
            xltype = (uint)XlTypes.xltypeNil;
            err = 0;
            if (value == null)
                return;

            if (value is double db)
            {
                num = db;
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is float fl)
            {
                num = fl;
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is decimal de)
            {
                num = Convert.ToDouble(de);
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is bool bl)
            {
                xltype = (uint)XlTypes.xltypeBool;
                w = bl ? -1 : 0;
            }
            else if (value is byte bt)
            {
                w = bt;
                xltype = (uint)XlTypes.xltypeInt;
            }
            else if (value is short st)
            {
                w = st;
                xltype = (uint)XlTypes.xltypeInt;
            }
            else if (value is int it)
            {
                w = it;
                xltype = (uint)XlTypes.xltypeInt;
            }
            else if (value is long lg)
            {
                num = lg;
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is string sr)
            {
                any = (char*)Marshal.AllocCoTaskMem((sr.Length + 1) * sizeof(char));
                char* p = (char*)any;
                p[0] = (char)sr.Length;
                for (var idx = 1; idx <= sr.Length; idx++)
                    p[idx] = sr[idx - 1];
                xltype = (uint)XlTypes.xltypeStr;
            }
            else if (value is DateTime dt)
            {
                num = dt.ToOADate();
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is object[] sa)
            {
                if (sa.Length == 0) return;
                array.rows = sa.Length;
                array.columns = 1;
                array.lparray = (XLOPER12*)Marshal.AllocHGlobal(sa.Length * sizeof(XLOPER12*));
                var row0 = sa.GetLowerBound(0);
                for (var row = row0; row <= sa.GetUpperBound(0); row++)
                {
                    var ele = array.lparray + row - row0;
                    ele->Init(sa[row]);
                }
                xltype = (uint)XlTypes.xltypeMulti;
            }
            else if (value is object[,] da)
            {
                if (da.Length == 0) return;
                array.rows = da.GetLength(0);
                array.columns = da.GetLength(1);
                var xxx = Marshal.SizeOf(typeof(XLOPER12));
                array.lparray = (XLOPER12*)Marshal.AllocHGlobal(array.rows * array.columns * sizeof(XLOPER12));
                var row0 = da.GetLowerBound(0);
                var col0 = da.GetLowerBound(1);
                for (var row = row0; row <= da.GetUpperBound(0); row++)
                    for (var col = col0; col <= da.GetUpperBound(1); col++)
                    {
                        var ele = array.lparray + (row - row0) * array.columns + col - col0;
                        ele->Init(da[row, col]);
                    }
                xltype = (uint)XlTypes.xltypeMulti;
            }
            xltype = (uint)((XlTypes)xltype | XlTypes.xlbitDLLFree);
        }

        private object ToObject()
        {
            var type = (XlTypes)xltype;
            switch (type)
            {
                case XlTypes.xltypeInt: return w;
                case XlTypes.xltypeNum: return num;
                case XlTypes.xltypeBool: return w != 0;
                case XlTypes.xltypeStr:
                    char* p = (char*)any;
                    var length = p[0];
                    if (length == 0)
                        return String.Empty;
                    var d = new char[length];
                    for (var idx = 1; idx <= length; idx++)
                        d[idx - 1] = p[idx];
                    return new string(d);
                case XlTypes.xltypeMulti:
                   var result = new object[array.rows, array.columns];
                    for (var row = 0; row < array.rows; row++)
                        for (var col = 0; col < array.columns; col++)
                        {
                            var ele = array.lparray + row * array.columns + col;
                            result[row, col] =ele->ToObject();
                        }
                   return result;
                case XlTypes.xltypeMissing:
                    return null;
            }
            return null;
        }
    }
}
