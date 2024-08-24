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

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public void InitObjectValue()
        {
            XLOPER12* p = (XLOPER12*)ObjectValue.ToPointer();
            p->Init(null, false);
        }

        public void FreeObjectValue()
        {
            XLOPER12* p = (XLOPER12*)ObjectValue.ToPointer();
            p->Dispose();
        }

        public IntPtr ObjectToIntPtr(object value)
        {
            XLOPER12* p = (XLOPER12*)ObjectValue.ToPointer();
            p->Init(value, true);
            return ObjectValue;
        }

        public IntPtr BooleanToIntPtr(bool value)
        {
            *((short*)ShortValue.ToPointer()) = value ? (short)1 : (short)0;
            return ShortValue;
        }

        public IntPtr DoubleToIntPtr(double value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr DateTimeToIntPtr(DateTime value)
        {
            *((double*)DoubleValue.ToPointer()) = value.ToOADate();
            return DoubleValue;
        }

        public IntPtr SingleToIntPtr(float value)
        {
            *((double*)DoubleValue.ToPointer()) = value;
            return DoubleValue;
        }

        public IntPtr Int32ToIntPtr(int value)
        {
            *((int*)IntValue.ToPointer()) = value;
            return IntValue;
        }

        public IntPtr UInt32ToIntPtr(uint value)
        {
            *((int*)IntValue.ToPointer()) =(int) value;
            return IntValue;
        }

        public IntPtr Int16ToIntPtr(short value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr UInt16ToIntPtr(ushort value)
        {
            return UInt32ToIntPtr(value);
        }

        public IntPtr ByteToIntPtr(byte value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr SByteToIntPtr(sbyte value)
        {
            *((short*)ShortValue.ToPointer()) = value;
            return ShortValue;
        }

        public IntPtr StringToIntPtr(string value)
        {
            const int SmallSize = 32768;
            var len = value?.Length ?? 0;
            if (len + 1 <= SmallSize)
            {
                if (StringValue == IntPtr.Zero)
                    StringValue = Marshal.AllocCoTaskMem(sizeof(char) * SmallSize);
                Copy(value, len, StringValue);
                return StringValue;
            }
            else
            {
                Marshal.FreeCoTaskMem(LargeStringValue);
                LargeStringValue = Marshal.AllocCoTaskMem(sizeof(char) * (len + 1));
                Copy(value, len, LargeStringValue);
                return LargeStringValue;
            }
        }

        public IntPtr DoubleArrayToIntPtr(double[] value)
        {
            if ((value?.Length ?? 0) == 0)
                return IntPtr.Zero;

            Marshal.FreeCoTaskMem(DoubleArrayValue);
            var len = value.Length;
            DoubleArrayValue = Marshal.AllocCoTaskMem(sizeof(int) * 2 + sizeof(double) * len);
            int* p = (int*)DoubleArrayValue.ToPointer();
            p[0] = 1;
            p[1] = len;

            double* d = (double*)&p[2];
            var col0 = value.GetLowerBound(0);
            for (var col = 0; col < len; col++)
                d[col] = value[col0 + col];
            return DoubleArrayValue;
        }

        public IntPtr DoubleMatrixToIntPtr(double[,] value)
        {
            if ((value?.Length ?? 0) == 0)
                return IntPtr.Zero;

            Marshal.FreeCoTaskMem(DoubleArrayValue);
            var rows = value.GetLength(0);
            var cols = value.GetLength(1);
            DoubleArrayValue = Marshal.AllocCoTaskMem(sizeof(int) * 2 + sizeof(double) * rows * cols);
            int* p = (int*)DoubleArrayValue.ToPointer();
            p[0] = rows;
            p[1] = cols;

            double* d = (double*)&p[2];
            var row0= value.GetLowerBound(0);
            var col0 = value.GetLowerBound(1);
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    d[row * cols + col] = value[row0 + row, col0 + col];
            return DoubleArrayValue;
        }

        public IntPtr DateTimeArrayToIntPtr(DateTime[] value)
        {
            if ((value?.Length ?? 0) == 0)
                return IntPtr.Zero;

            var cols = value.Length;
            var cells = new double[cols];
            var col0 = value.GetLowerBound(0);
            for (var col = 0; col < cols; col++)
                cells[col] = value[col0 + col].ToOADate();
            return DoubleArrayToIntPtr(cells);
        }

        public IntPtr DateTimeMatrixToIntPtr(DateTime[,] value)
        {
            if ((value?.Length ?? 0) == 0)
                return IntPtr.Zero;

            var rows = value.GetLength(0);
            var cols = value.GetLength(1);
            var cells = new double[rows, cols];
            var row0 = value.GetLowerBound(0);
            var col0 = value.GetLowerBound(1);
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    cells[row, col] = value[row0 + row, col0 + col].ToOADate();
            return DoubleMatrixToIntPtr(cells);
        }

        public IntPtr Int32ArrayToIntPtr(int[] value)
        {
            if ((value?.Length ?? 0) == 0)
                return IntPtr.Zero;

            var cols = value.Length;
            var cells = new double[cols];
            var col0 = value.GetLowerBound(0);
            for (var col = 0; col < cols; col++)
                cells[col] = value[col0 + col];
            return DoubleArrayToIntPtr(cells);
        }

        public IntPtr Int32MatrixToIntPtr(int[,] value)
        {
            if ((value?.Length ?? 0) == 0)
                return IntPtr.Zero;
            var rows = value.GetLength(0);
            var cols = value.GetLength(1);
            var cells = new double[rows, cols];
            var row0 = value.GetLowerBound(0);
            var col0 = value.GetLowerBound(1);
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    cells[row, col] = value[row0 + row, col0 + col];
            return DoubleMatrixToIntPtr(cells);
        }

        public IntPtr ObjectArrayToIntPtr(object[] value)
        {
            var x = (XLOPER12 *) ObjectValue.ToPointer();
            x->Init(value, true);
            return ObjectValue;
        }

        public IntPtr ObjectMatrixToIntPtr(object[,] value)
        {
            var x = (XLOPER12*)ObjectValue.ToPointer();
            x->Init(value, true);
            return ObjectValue;
        }

        public IntPtr StringArrayToIntPtr(string[] value)
        {
            var x = (XLOPER12*)ObjectValue.ToPointer();
            x->Init(value, true);
            return ObjectValue;
        }

        public IntPtr StringMatrixToIntPtr(string[,] value)
        {
            var x = (XLOPER12*)ObjectValue.ToPointer();
            x->Init(value, true);
            return ObjectValue;
        }

        private static void Copy(string source, int length, IntPtr target)
        {
            char* p = (char*)target.ToPointer();
            p[length] = '\0';
            for (var idx = 0; idx < length; idx++)
                p[idx] = source[idx];
        }

        private static readonly Dictionary<Type, MethodInfo> OutgoingConverters
            = new Dictionary<Type, MethodInfo>()
            {
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(BooleanToIntPtr)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(DoubleToIntPtr)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(DateTimeToIntPtr)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(SingleToIntPtr)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(Int32ToIntPtr)) },
                { typeof(uint), typeof(XlMarshalContext).GetMethod(nameof(UInt32ToIntPtr)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(Int16ToIntPtr)) },
                { typeof(ushort), typeof(XlMarshalContext).GetMethod(nameof(UInt16ToIntPtr)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(ByteToIntPtr)) },
                { typeof(sbyte), typeof(XlMarshalContext).GetMethod(nameof(SByteToIntPtr)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(StringToIntPtr)) },
                { typeof(double[]), typeof(XlMarshalContext).GetMethod(nameof(DoubleArrayToIntPtr)) },
                { typeof(double[,]), typeof(XlMarshalContext).GetMethod(nameof(DoubleMatrixToIntPtr)) },
                { typeof(int[]), typeof(XlMarshalContext).GetMethod(nameof(Int32ArrayToIntPtr)) },
                { typeof(int[,]), typeof(XlMarshalContext).GetMethod(nameof(Int32MatrixToIntPtr)) },
                { typeof(DateTime[]), typeof(XlMarshalContext).GetMethod(nameof(DateTimeArrayToIntPtr)) },
                { typeof(DateTime[,]), typeof(XlMarshalContext).GetMethod(nameof(DateTimeMatrixToIntPtr)) },
                { typeof(string[]), typeof(XlMarshalContext).GetMethod(nameof(StringArrayToIntPtr)) },
                { typeof(string[,]), typeof(XlMarshalContext).GetMethod(nameof(StringMatrixToIntPtr)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(ObjectToIntPtr)) },
                { typeof(object[]), typeof(XlMarshalContext).GetMethod(nameof(ObjectArrayToIntPtr)) },
                { typeof(object[,]), typeof(XlMarshalContext).GetMethod(nameof(ObjectMatrixToIntPtr)) }
            };

        public static MethodInfo OutgoingConverter(Type returnType) =>
            OutgoingConverters.TryGetValue(returnType, out var value) ? value : OutgoingConverters[(typeof(object))];
    }
}
