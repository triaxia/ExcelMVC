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
using System.IO;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class XlMarshalExceptionHandler
    {
        public static event EventHandler<ErrorEventArgs> Failed;
        public static Func<Exception, object> ExceptionToFunctionResult { get; set; }

        public static object HandleException(Exception ex)
        {
            try
            {
                Failed?.Invoke(null, new ErrorEventArgs(ex));
                return ExceptionToFunctionResult?.Invoke(ex) ?? ExcelError.ExcelErrorValue;
            }
            catch (Exception fatal)
            {
                return $"{fatal}";// ExcelError.ExcelErrorValue;
            }
        }

        public static MethodInfo HandlerMethod =>
            typeof(XlMarshalExceptionHandler).GetMethod(nameof(XlMarshalExceptionHandler.HandleException)
                , BindingFlags.Static | BindingFlags.Public);
    }

    public unsafe partial class XlMarshalContext
    {
        public IntPtr ObjectToIntPtrOnException(object value)
        {
            XLOPER12* p = (XLOPER12*)ObjectValue.ToPointer();
            p->Init(value, true);
            return ObjectValue;
        }

        public IntPtr IntToIntPtrOnException(object value)
        {
            return IntPtr.Zero;
        }

        public IntPtr StringToIntPtrOnException(object value)
        {
            return StringToIntPtr($"{value}");
        }

        private static readonly Dictionary<Type, MethodInfo> ExceptionConverters
            = new Dictionary<Type, MethodInfo>()
            {
                /*
                { typeof(bool), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(double), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(DateTime), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(float), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(uint), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(short), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(ushort), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(byte), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(sbyte), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(double[]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(double[,]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(int[]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(int[,]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(DateTime[]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(DateTime[,]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(string[]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(string[,]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(ExceptionObjectToIntPtr)) },
                { typeof(object[]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) },
                { typeof(object[,]), typeof(XlMarshalContext).GetMethod(nameof(ExceptionAnyToIntPtr)) }
                */
                { typeof(string), typeof(XlMarshalContext).GetMethod(nameof(StringToIntPtrOnException)) },
                { typeof(int), typeof(XlMarshalContext).GetMethod(nameof(IntToIntPtrOnException)) },
                { typeof(object), typeof(XlMarshalContext).GetMethod(nameof(ObjectToIntPtrOnException)) },
            };

        public static MethodInfo ExceptionConverter(Type returnType) =>
            ExceptionConverters.TryGetValue(returnType, out var value) ? value : ExceptionConverters[(typeof(int))];
    }
}
