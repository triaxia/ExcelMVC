/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor,
Boston, MA 02110-1301 USA.
*/

using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Sequential)]
    unsafe public struct CallStatus
    {
        public XLOPER12* result;
        public int status;
    }

    [StructLayout(LayoutKind.Sequential)]
    unsafe public struct XLBigData
    {
        public IntPtr data;
        public long cbCount;
    }

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
        [FieldOffset(0)] public char* str;
        [FieldOffset(0)] public XLBigData bigdata;
        [FieldOffset(0)] public XLArray array;
        [FieldOffset(24)] public uint xltype;

        public static XLOPER12 FromObject(object value)
        {
            return new XLOPER12(value);
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
            var type = PeelOfType((XlTypes)xltype);
            if (type == XlTypes.xltypeStr && str != null)
            {
                str = null;
                Marshal.FreeCoTaskMem((IntPtr)str);
            }
            if (type == XlTypes.xltypeMulti && array.lparray != null)
            {
                array.lparray = null;
                Marshal.FreeCoTaskMem((IntPtr)array.lparray);
            }
            xltype = (uint) XlTypes.xltypeNil;
        }

        public XLOPER12(object value)
        {
            num = 0;
            w = 0;
            str = null;
            array = new XLArray
            {
                rows = 0,
                columns = 0,
                lparray = null
            };
            bigdata = new XLBigData
            {
                data = IntPtr.Zero,
                cbCount = 0,
            };
            xltype = (uint)XlTypes.xltypeNil;
            err = 0;
            Init(value, false);
        }

        public void Init(object value, bool dispose)
        {
            if (dispose) Dispose();
            num = 0;
            w = 0;
            str = null;
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
            else if (value is ushort ust)
            {
                w = ust;
                xltype = (uint)XlTypes.xltypeInt;
            }
            else if (value is int it)
            {
                w = it;
                xltype = (uint)XlTypes.xltypeInt;
            }
            else if (value is uint uit)
            {
                w = (int)uit;
                xltype = (uint)XlTypes.xltypeInt;
            }
            else if (value is long ln)
            {
                Init(ln.ToString(), true);
            }
            else if (value is ulong uln)
            {
                Init(uln.ToString(), true);
            }
            else if (value is string sr)
            {
                str = (char*)Marshal.AllocCoTaskMem((sr.Length + 1) * sizeof(char));
                str[0] = (char)sr.Length;
                for (var idx = 1; idx <= sr.Length; idx++)
                    str[idx] = sr[idx - 1];
                str[sr.Length + 1] = (char)0;
                xltype = (uint)XlTypes.xltypeStr;
            }
            else if (value is DateTime dt)
            {
                num = dt.ToOADate();
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is object[] sa)
            {
                if (sa.Length == 0)
                {
                    Init("", true);
                }
                else
                {
                    array.rows = sa.Length > 0 ? 1 : 0;
                    array.columns = sa.Length;
                    array.lparray = (XLOPER12*)Marshal.AllocCoTaskMem(sa.Length * sizeof(XLOPER12));
                    var col0 = sa.GetLowerBound(0);
                    var col1 = sa.GetUpperBound(0);
                    for (var col = col0; col <= col1; col++)
                    {
                        var ele = array.lparray + col - col0;
                        ele->Init(sa[col], false);
                    }
                    xltype = (uint)XlTypes.xltypeMulti;
                }
            }
            else if (value is object[,] da)
            {
                if (da.Length == 0)
                {
                    Init("", true);
                }
                else
                {
                    array.rows = da.GetLength(0);
                    array.columns = da.GetLength(1);
                    array.lparray = (XLOPER12*)Marshal.AllocCoTaskMem(array.rows * array.columns * sizeof(XLOPER12));
                    var row0 = da.GetLowerBound(0);
                    var row1 = da.GetUpperBound(0);
                    var col0 = da.GetLowerBound(1);
                    var col1 = da.GetUpperBound(1);
                    for (var row = row0; row <= row1; row++)
                        for (var col = col0; col <= col1; col++)
                        {
                            var ele = array.lparray + (row - row0) * array.columns + col - col0;
                            ele->Init(da[row, col], false);
                        }
                    xltype = (uint)XlTypes.xltypeMulti;
                }
            }
            else if (value is ExcelError xle)
            {
                xltype = (uint)XlTypes.xltypeErr;
                err = (int)xle;
            }
            else if (value is ExcelMissing)
            {
                xltype = (uint)XlTypes.xltypeMissing;
            }
            else if (value is ExcelEmpty)
            {
                xltype = (uint)XlTypes.xltypeNil;
            }
            else if (value is IntPtr)
            {
                bigdata.data = (IntPtr)value;
                bigdata.cbCount = 0;
                xltype = (uint)XlTypes.xltypeBigData;
            }
            xltype = (uint) ((XlTypes)xltype | XlTypes.xlbitDLLFree); 
        }

        public object ToObject()
        {
            var type = PeelOfType((XlTypes)xltype);
            switch (type)
            {
                case XlTypes.xltypeInt: return w;
                case XlTypes.xltypeNum: return num;
                case XlTypes.xltypeBool: return w != 0;
                case XlTypes.xltypeStr:
                    var length = str[0];
                    if (length == 0)
                        return string.Empty;
                    var d = new char[length];
                    for (var idx = 1; idx <= length; idx++)
                        d[idx - 1] = str[idx];
                    return new string(d);
                case XlTypes.xltypeMulti:
                    if (array.rows == 0 || array.columns == 0)
                        return new object[] { };
                    var result = new object[array.rows, array.columns];
                    for (var row = 0; row < array.rows; row++)
                        for (var col = 0; col < array.columns; col++)
                        {
                            var ele = array.lparray + row * array.columns + col;
                            result[row, col] = ele->ToObject();
                        }
                    return result;
                case XlTypes.xltypeNil:
                    return ExcelEmpty.Value;
                case XlTypes.xltypeMissing:
                    return ExcelMissing.Value;
                case XlTypes.xltypeErr:
                    return (ExcelError)err;
            }
            return null;
        }

        public object[] ToObjectArray()
        {
            var type = PeelOfType((XlTypes)xltype);
            switch (type)
            {
                case XlTypes.xltypeInt: return new object[] { w };
                case XlTypes.xltypeNum: return new object[] { num };
                case XlTypes.xltypeBool: return new object[] { w != 0 };
                case XlTypes.xltypeStr:
                    var length = str[0];
                    if (length == 0)
                        return new object[] { String.Empty };
                    var d = new char[length];
                    for (var idx = 1; idx <= length; idx++)
                        d[idx - 1] = str[idx];
                    return new object[] { new string(d) };
                case XlTypes.xltypeMulti:
                    if (array.rows == 0 || array.columns == 0)
                        return new object[] { };
                    var result = new object[array.rows * array.columns];
                    for (var row = 0; row < array.rows; row++)
                        for (var col = 0; col < array.columns; col++)
                        {
                            var ele = array.lparray + row * array.columns + col;
                            result[row * array.columns + col] = ele->ToObject();
                        }
                    return result;
                case XlTypes.xltypeNil:
                    return new object[] { ExcelEmpty.Value };
                case XlTypes.xltypeMissing:
                    return new object[] { ExcelMissing.Value };
                case XlTypes.xltypeErr:
                    return new object[] { (ExcelError)err };
            }
            return new object[] { };
        }

        public object[,] ToObjectMatrix()
        {
            var type = PeelOfType((XlTypes)xltype);
            switch (type)
            {
                case XlTypes.xltypeInt: return new object[,] { { w } };
                case XlTypes.xltypeNum: return new object[,] { { num } };
                case XlTypes.xltypeBool: return new object[,] { { w != 0 } };
                case XlTypes.xltypeStr:
                    var length = str[0];
                    if (length == 0)
                        return new object[,] { { String.Empty } };
                    var d = new char[length];
                    for (var idx = 1; idx <= length; idx++)
                        d[idx - 1] = str[idx];
                    return new object[,] { { new string(d) } };
                case XlTypes.xltypeMulti:
                    if (array.rows == 0 || array.columns == 0)
                        return new object[,] { };
                    var result = new object[array.rows, array.columns];
                    for (var row = 0; row < array.rows; row++)
                        for (var col = 0; col < array.columns; col++)
                        {
                            var ele = array.lparray + row * array.columns + col;
                            result[row, col] = ele->ToObject();
                        }
                    return result;
                case XlTypes.xltypeNil:
                    return new object[,] { { ExcelEmpty.Value } };
                case XlTypes.xltypeMissing:
                    return new object[,] { { ExcelMissing.Value } };
                case XlTypes.xltypeErr:
                    return new object[,] { { (ExcelError)err } };
            }
            return new object[,] { };
        }
        private static XlTypes PeelOfType(XlTypes type)
        {
            return (XlTypes)type & ~XlTypes.xlbitDLLFree & ~XlTypes.xlbitXLFree;
        }
    }
}
