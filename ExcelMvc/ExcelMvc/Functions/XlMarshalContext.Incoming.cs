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
using System.Linq;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        public static IntPtr IntPtrToIntPtr(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            return value;
        }

        public static bool IntPtrToBoolean(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<bool>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? false : *(short*)value.ToPointer() != 0;
        }

        public static bool? IntPtrToBooleanNullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<bool?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(bool?) : *(short*)value.ToPointer() != 0;
        }

        public static double IntPtrToDouble(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? 0 : *(double*)value.ToPointer();
        }

        public static double? IntPtrToDoubleNullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(double?) : *(double*)value.ToPointer();
        }

        public static DateTime IntPtrToDateTime(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double>(value, parameter, isOptional, out var output))
                return DateTime.FromOADate(output);
            return value == IntPtr.Zero ? DateTime.FromOADate(0) : DateTime.FromOADate(*(double*)value.ToPointer());
        }

        public static DateTime? IntPtrToDateTimeNullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double?>(value, parameter, isOptional, out var output))
                return output.HasValue ? DateTime.FromOADate(output.Value) : default(DateTime?);
            return value == IntPtr.Zero ? default(DateTime?) : DateTime.FromOADate(*(double*)value.ToPointer());
        }

        public static float IntPtrToSingle(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<float>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? 0 : (float)*(double*)value.ToPointer();
        }

        public static float? IntPtrToSingleNullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<float?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(float?) : (float)*(double*)value.ToPointer();
        }

        public static int IntPtrToInt32(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<int>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? 0 : *(int*)value.ToPointer();
        }

        public static int? IntPtrToInt32Nullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<int?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(int?) : *(int*)value.ToPointer();
        }

        public static uint IntPtrToUInt32(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<uint>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? 0 : (uint)*(int*)value.ToPointer();
        }

        public static uint? IntPtrToUInt32Nullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<uint?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(uint?) : (uint)*(int*)value.ToPointer();
        }

        public static short IntPtrToInt16(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<short>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? (short)0 : *(short*)value.ToPointer();
        }

        public static short? IntPtrToInt16Nullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<short?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(short?) : *(short*)value.ToPointer();
        }

        public static ushort IntPtrToUInt16(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<ushort>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? (ushort)0 : (ushort)*(short*)value.ToPointer();
        }

        public static ushort? IntPtrToUInt16Nullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<ushort?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(ushort?) : (ushort)*(short*)value.ToPointer();
        }

        public static byte IntPtrToByte(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<byte>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? (byte)0 : (byte)*(short*)value.ToPointer();
        }

        public static byte? IntPtrToByteNullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<byte?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(byte?) : (byte)*(short*)value.ToPointer();
        }

        public static sbyte IntPtrToSByte(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<sbyte>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? (sbyte)0 : (sbyte)*(short*)value.ToPointer();
        }

        public static sbyte? IntPtrToSByteNullable(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<sbyte?>(value, parameter, isOptional, out var output))
                return output;
            return value == IntPtr.Zero ? default(sbyte?) : (sbyte)*(short*)value.ToPointer();
        }

        public static string IntPtrToString(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<string>(value, parameter, isOptional, out var output))
                return output;

            if (value == IntPtr.Zero)
                return string.Empty;

            char* p = (char*)value.ToPointer();
            return new string(p);
        }

        public static double[] IntPtrToDoubleArray(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double[]>(value, parameter, isOptional, out var output))
                return output;

            if (value == IntPtr.Zero)
                return new double[] { };

            int* p = (int*)value.ToPointer();
            var rows = p[0];
            var cols = p[1];
            if (rows == 0 || cols == 0)
                return new double[] { };
            var len = rows * cols;
            var result = new double[len];
            double* x = (double*)&p[2];
            for (var i = 0; i < len; i++)
                result[i] = x[i];
            return result;
        }

        public static double[,] IntPtrToDoubleMatrix(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double[,]>(value, parameter, isOptional, out var output))
                return output;

            if (value == IntPtr.Zero)
                return new double[,] { };

            int* p = (int*)value.ToPointer();
            var rows = p[0];
            var cols = p[1];
            if (rows == 0 || cols == 0)
                return new double[,] { };
            var len = rows * cols;
            var result = new double[rows, cols];
            double* x = (double*)&p[2];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    result[row, col] = x[row * cols + col];
            return result;
        }

        public static DateTime[] IntPtrToDateTimeArray(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double[]>(value, parameter, isOptional, out var output))
                return ToDateTime(output);

            if (value == IntPtr.Zero)
                return new DateTime[] { };

            return ToDateTime(IntPtrToDoubleArray(value, parameter, isOptional));
        }

        public static DateTime[,] IntPtrToDateTimeMatrix(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<double[,]>(value, parameter, isOptional, out var output))
                return ToDateTime(output);

            if (value == IntPtr.Zero)
                return new DateTime[,] { };
            return ToDateTime(IntPtrToDoubleMatrix(value, parameter, isOptional));
        }

        public static int[] IntPtrToInt32Array(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<int[]>(value, parameter, isOptional, out var output))
                return output;

            if (value == IntPtr.Zero)
                return new int[] { };

            var cells = IntPtrToDoubleArray(value, parameter, isOptional);
            if (cells.Length == 0)
                return new int[] { };

            var result = new int[cells.Length];
            for (var i = 0; i < cells.Length; i++)
                result[i] = (int)cells[i];
            return result;
        }

        public static int[,] IntPtrToInt32Matrix(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (TryGetOptionalValue<int[,]>(value, parameter, isOptional, out var output))
                return output;

            if (value == IntPtr.Zero)
                return new int[,] { };

            var cells = IntPtrToDoubleMatrix(value, parameter, isOptional);
            if ((cells?.Length ?? 0) == 0)
                return new int[,] { };
            var rows = cells.GetLength(0);
            var cols = cells.GetLength(1);
            var result = new int[rows, cols];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    result[row, col] = (int)cells[row, col];
            return result;
        }

        public static object IntPtrToObject(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (value == IntPtr.Zero)
                return null;
            XLOPER12* p = (XLOPER12*)value.ToPointer();
            return p->ToObject();
        }

        public static object[] IntPtrToObjectArray(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (value == IntPtr.Zero)
                return new object[] { };
            XLOPER12* p = (XLOPER12*)value.ToPointer();
            return p->ToObjectArray();
        }

        public static object[,] IntPtrToObjectMatrix(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (value == IntPtr.Zero)
                return new object[,] { };
            XLOPER12* p = (XLOPER12*)value.ToPointer();
            return p->ToObjectMatrix();
        }

        public static string[] IntPtrToStringArray(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (value == IntPtr.Zero)
                return new string[] { };
            XLOPER12* p = (XLOPER12*)value.ToPointer();
            return p->ToObjectArray().Select(x => $"{x}").ToArray();
        }

        public static string[,] IntPtrToStringMatrix(IntPtr value, ParameterInfo parameter, bool isOptional)
        {
            if (value == IntPtr.Zero)
                return new string[,] { };
            XLOPER12* p = (XLOPER12*)value.ToPointer();
            var cells = p->ToObjectMatrix();
            if ((cells?.Length ?? 0) == 0)
                return new string[,] { };
            var rows = cells.GetLength(0);
            var cols = cells.GetLength(1);
            var result = new string[rows, cols];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    result[row, col] = $"{cells[row, col]}";
            return result;
        }

        public static bool TryGetOptionalValue<TValue>(IntPtr value, ParameterInfo parameter, bool isOptional, out TValue result)
        {
            result = default;
            if (!isOptional) return false;

            var objValue = IntPtrToObject(value, parameter, isOptional);
            if (objValue is ExcelMissing)
                objValue = parameter.DefaultValue == DBNull.Value ? default : parameter.DefaultValue;

            if (typeof(TValue).IsValueType)
                result = objValue == null ? default : ChangeType<TValue>(objValue);
            else
                result = (TValue)objValue;

            return true;
        }

        public static DateTime[] ToDateTime(double[] cells)
        {
            if ((cells?.Length ?? 0) == 0)
                return new DateTime[] { };

            var result = new DateTime[cells.Length];
            for (var i = 0; i < cells.Length; i++)
                result[i] = DateTime.FromOADate(cells[i]);
            return result;
        }

        public static DateTime[,] ToDateTime(double[,] cells)
        {
            var rows = cells?.GetLength(0) ?? 0;
            var cols = cells?.GetLength(1) ?? 0;
            if (rows == 0 || cols == 0)
                return new DateTime[,] { };
            var result = new DateTime[rows, cols];
            for (var row = 0; row < rows; row++)
                for (var col = 0; col < cols; col++)
                    result[row, col] = DateTime.FromOADate(cells[row, col]);
            return result;
        }

        private static T ChangeType<T>(object value)
        {
            var t = typeof(T);

            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                if (value == null) return default;
                t = Nullable.GetUnderlyingType(t);
            }
            return (T)Convert.ChangeType(value, t);
        }

        private static readonly Dictionary<Type, MethodInfo> IncomingConverters
            = new Dictionary<Type, MethodInfo>()
            {
                { typeof(IntPtr), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToIntPtr)) },
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToBoolean)) },
                { typeof(bool?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToBooleanNullable)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDouble)) },
                { typeof(double?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDoubleNullable)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTime)) },
                { typeof(DateTime?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTimeNullable)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToSingle)) },
                { typeof(float?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToSingleNullable)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt32)) },
                { typeof(int?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt32Nullable)) },
                { typeof(uint), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUInt32)) },
                { typeof(uint?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUInt32Nullable)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt16)) },
                { typeof(short?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt16Nullable)) },
                { typeof(ushort), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUInt16)) },
                { typeof(ushort?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToUInt16Nullable)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToByte)) },
                { typeof(byte?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToByteNullable)) },
                { typeof(sbyte), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToSByte)) },
                { typeof(sbyte?), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToSByteNullable)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToString)) },
                { typeof(double[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDoubleArray)) },
                { typeof(double[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDoubleMatrix)) },
                { typeof(int[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt32Array)) },
                { typeof(int[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToInt32Matrix)) },
                { typeof(DateTime[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTimeArray)) },
                { typeof(DateTime[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToDateTimeMatrix)) },
                { typeof(string[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToStringArray)) },
                { typeof(string[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToStringMatrix)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToObject)) },
                { typeof(object[]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToObjectArray)) },
                { typeof(object[,]), typeof(XlMarshalContext).GetMethod(nameof(IntPtrToObjectMatrix)) }
            };

        public static MethodInfo IncomingConverter(Type parameterType) =>
            IncomingConverters.TryGetValue(parameterType, out var value) ? value : IncomingConverters[(typeof(object))];
    }
}
