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

using ExcelMvc.Diagnostics;
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
            if (xltype == (uint)XlTypes.xltypeStr && any != null)
                Marshal.FreeCoTaskMem((IntPtr)any);
            if (xltype == (uint)XlTypes.xltypeMulti && array.lparray != null)
                Marshal.FreeCoTaskMem((IntPtr)array.lparray);
        }

        public XLOPER12(object value)
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
            Init(value, false);
        }

        public void Init(object value, bool dispose)
        {
            if (dispose) Dispose();
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
            else if (value is long lg)
            {
                num = lg;
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is ulong ulg)
            {
                num = ulg;
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is string sr)
            {
                any = (char*)Marshal.AllocCoTaskMem((sr.Length + 1) * sizeof(char));
                char* p = (char*)any;
                p[0] = (char)sr.Length;
                for (var idx = 1; idx <= sr.Length; idx++)
                    p[idx] = sr[idx - 1];
                p[sr.Length + 1] = (char)0;
                xltype = (uint)XlTypes.xltypeStr;
            }
            else if (value is DateTime dt)
            {
                num = dt.ToOADate();
                xltype = (uint)XlTypes.xltypeNum;
            }
            else if (value is object[] sa)
            {
                array.rows = sa.Length > 0 ? 1 : 0;
                array.columns = sa.Length;
                if (array.rows != 0 && array.columns != 0)
                {
                    array.lparray = (XLOPER12*)Marshal.AllocCoTaskMem(sa.Length * sizeof(XLOPER12));
                    var col0 = sa.GetLowerBound(0);
                    var colx = sa.GetUpperBound(0);
                    for (var col = col0; col <= colx; col++)
                    {
                        var ele = array.lparray + col - col0;
                        ele->Init(sa[col], false);
                    }
                }
                xltype = (uint)XlTypes.xltypeMulti;
            }
            else if (value is object[,] da)
            {
                array.rows = da.GetLength(0);
                array.columns = da.GetLength(1);
                if (array.rows != 0 && array.columns != 0)
                {
                    array.lparray = (XLOPER12*)Marshal.AllocCoTaskMem(array.rows * array.columns * sizeof(XLOPER12));
                    var row0 = da.GetLowerBound(0);
                    var rowx = da.GetUpperBound(0);
                    var col0 = da.GetLowerBound(1);
                    var colx = da.GetUpperBound(1);
                    for (var row = row0; row <= rowx; row++)
                        for (var col = col0; col <= colx; col++)
                        {
                            var ele = array.lparray + (row - row0) * array.columns + col - col0;
                            ele->Init(da[row, col], false);
                        }
                }
                xltype = (uint)XlTypes.xltypeMulti;
            }
            else if (value is XlError xle)
            {
                xltype = (uint)XlTypes.xltypeErr;
                err = (int)XlErrorFactory.ObjectToType(xle);
            }
            else if (value is XlMissing)
            {
                xltype = (uint)XlTypes.xltypeMissing;
            }
            else if (value is XlEmpty)
            {
                xltype = (uint)XlTypes.xltypeNil;
            }
        }

        public object ToObject()
        {
            var type = (XlTypes)xltype & ~XlTypes.xlbitDLLFree & ~XlTypes.xlbitXLFree;
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
                    return XlEmpty.Instance;
                case XlTypes.xltypeMissing:
                    return XlMissing.Instance;
                case XlTypes.xltypeErr:
                    return XlErrorFactory.TypeToObject((XlErrors)err);
            }
            return null;
        }

        public object[] ToObjectArray()
        {
            var type = (XlTypes)xltype & ~XlTypes.xlbitDLLFree & ~XlTypes.xlbitXLFree;
            switch (type)
            {
                case XlTypes.xltypeInt: return new object[] { w };
                case XlTypes.xltypeNum: return new object[] { num };
                case XlTypes.xltypeBool: return new object[] { w != 0 };
                case XlTypes.xltypeStr:
                    char* p = (char*)any;
                    var length = p[0];
                    if (length == 0)
                        return new object[] { String.Empty };
                    var d = new char[length];
                    for (var idx = 1; idx <= length; idx++)
                        d[idx - 1] = p[idx];
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
                    return new object[] { XlEmpty.Instance };
                case XlTypes.xltypeMissing:
                    return new object[] { XlMissing.Instance };
                case XlTypes.xltypeErr:
                    return new object[] { XlErrorFactory.TypeToObject((XlErrors)err) };
            }
            return new object[] { };
        }

        public object[,] ToObjectMatrix()
        {
            var type = (XlTypes)xltype & ~XlTypes.xlbitDLLFree & ~XlTypes.xlbitXLFree;
            switch (type)
            {
                case XlTypes.xltypeInt: return new object[,] { { w } };
                case XlTypes.xltypeNum: return new object[,] { { num } };
                case XlTypes.xltypeBool: return new object[,] { { w != 0 } };
                case XlTypes.xltypeStr:
                    char* p = (char*)any;
                    var length = p[0];
                    if (length == 0)
                        return new object[,] { { String.Empty } };
                    var d = new char[length];
                    for (var idx = 1; idx <= length; idx++)
                        d[idx - 1] = p[idx];
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
                    return new object[,] { { XlEmpty.Instance } };
                case XlTypes.xltypeMissing:
                    return new object[,] { { XlMissing.Instance } };
                case XlTypes.xltypeErr:
                    return new object[,] { { XlErrorFactory.TypeToObject((XlErrors)err) } };
            }
            return new object[,] { };
        }
    }
}
