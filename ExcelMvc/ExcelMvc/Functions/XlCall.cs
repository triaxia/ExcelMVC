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

using ExcelMvc.Rtd;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class XlCall
    {
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "RegisterFunction")]
        public static extern IntPtr RegisterFunction64(IntPtr function);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "RegisterFunction")]
        public static extern IntPtr RegisterFunction32(IntPtr function);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "AsyncReturn")]
        public static extern IntPtr AsyncReturn64(IntPtr handle, IntPtr result);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "AsyncReturn")]
        public static extern IntPtr AsyncReturn32(IntPtr handle, IntPtr result);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "xlAutoFree12")]
        public static extern IntPtr xlAutoFree64(IntPtr handle);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "xlAutoFree12")]
        public static extern IntPtr xlAutoFree32(IntPtr handle);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "RtdCall")]
        public static extern IntPtr RtdCall64(IntPtr args);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "RtdCall")]
        public static extern IntPtr RtdCall32(IntPtr args);

        public static void RegisterFunction(Function function)
        {
            using (var pFunction = new StructIntPtr<Function>(ref function))
            {
                if (Environment.Is64BitProcess)
                   RegisterFunction64(pFunction.Ptr);
                else
                   RegisterFunction32(pFunction.Ptr);
            }
        }

        public static void AsyncReturn(IntPtr handle, IntPtr result)
        {
            if (Environment.Is64BitProcess)
                xlAutoFree64(AsyncReturn64(handle, result));
            else
                xlAutoFree32(AsyncReturn32(handle, result));
        }

        public static void SetAsyncResult(IntPtr handle, object result)
        {
            var outcome = XLOPER12.FromObject(result);
            try
            {
                using (var ptr = new StructIntPtr<XLOPER12>(ref outcome))
                    AsyncReturn(handle, ptr.Ptr);
            }
            finally
            {
                outcome.Dispose();
            }
        }

        public unsafe static object CallRtd(Type implType, Func<IRtdServerImpl> implFactory
            , string arg0, params string[] args)
        {
            using (var reg = new RtdRegistry(implType, implFactory))
            {
                var arguments = new string[] { reg.ProgId, "", arg0 }
                    .Concat(args)
                    .Select((x, idx) => new FunctionArgument($"p{idx}", x))
                    .ToArray();
                var fArgs = new FunctionArguments(arguments);
                IntPtr ptr = IntPtr.Zero;
                using (var pArgs = new StructIntPtr<FunctionArguments>(ref fArgs))
                {
                    if (Environment.Is64BitProcess)
                        ptr = RtdCall64(pArgs.Ptr);
                    else
                        ptr = RtdCall32(pArgs.Ptr);
                }
                var result = (XLOPER12*)ptr.ToPointer();
                return result == null ? null : result->ToObject();
            }
        }
    }
}
