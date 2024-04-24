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
using ExcelMvc.Rtd;
using ExcelMvc.Views;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public class MessageEventArgs : EventArgs
    {
        public string Message { get; }
        public MessageEventArgs(string message)
            => Message = message;
    }

    public static class XlCall
    {
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "RegisterFunction")]
        internal static extern IntPtr RegisterFunction64(IntPtr function);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "RegisterFunction")]
        internal static extern IntPtr RegisterFunction32(IntPtr function);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "AsyncReturn")]
        internal static extern IntPtr AsyncReturn64(IntPtr handle, IntPtr result);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "AsyncReturn")]
        internal static extern IntPtr AsyncReturn32(IntPtr handle, IntPtr result);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "xlAutoFree12")]
        internal static extern IntPtr xlAutoFree64(IntPtr handle);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "xlAutoFree12")]
        internal static extern IntPtr xlAutoFree32(IntPtr handle);
        [DllImport("ExcelMvc.Addin.x64.xll", EntryPoint = "RtdCall")]
        internal static extern IntPtr RtdCall64(IntPtr args);
        [DllImport("ExcelMvc.Addin.x86.xll", EntryPoint = "RtdCall")]
        internal static extern IntPtr RtdCall32(IntPtr args);

        internal static void RegisterFunction(Function function)
        {
            using (var pFunction = new StructIntPtr<Function>(ref function))
            {
                if (Environment.Is64BitProcess)
                   RegisterFunction64(pFunction.Ptr);
                else
                   RegisterFunction32(pFunction.Ptr);
            }
        }

        internal static void AsyncReturn(IntPtr handle, IntPtr result)
        {
            if (Environment.Is64BitProcess)
                xlAutoFree64(AsyncReturn64(handle, result));
            else
                xlAutoFree32(AsyncReturn32(handle, result));
        }

        /// <summary>
        /// Occurs whenever errors are encountered.
        /// </summary>
        public static event EventHandler<ErrorEventArgs> Failed;

        /// <summary>
        /// Raises <see cref="Failed"/> event.
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="sender"></param>
        public static void OnFailed(Exception ex, object sender = null)
        {
            Messages.Instance.AddErrorLine(ex);
            Failed?.Invoke(sender, new ErrorEventArgs(ex));
        }

        /// <summary>
        /// Occurs whenever messages are posted. 
        /// </summary>
        public static event EventHandler<MessageEventArgs> Posted;

        /// <summary>
        /// Raises <see cref="Posted"/> event.
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="sender"></param>
        public static void OnPosted(string message, object sender = null)
        {
            Messages.Instance.AddInfoLine(message);
            Posted?.Invoke(sender, new MessageEventArgs(message));
        }

        /// <summary>
        /// Gets the Excel caller range
        /// </summary>
        /// <returns></returns>
        public static ExcelReference GetCaller()
        {
            dynamic caller = App.Instance.Underlying.Caller;
            return caller is Microsoft.Office.Interop.Excel.Range range ? new ExcelReference(range)
                : new ExcelReference();
        }

        /// <summary>
        /// Gets the Async handle from the specified XL handle (XLOPER12)
        /// </summary>
        /// <param name="handle"></param>
        /// <returns></returns>
        public static IntPtr GetAsyncHandle(IntPtr handle)
        {
            unsafe
            {
                var p = (XLOPER12*)handle.ToPointer();
                return p->bigdata.data;
            }
        }

        /// <summary>
        /// Sets the asynchronous result.
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="result"></param>
        public static void SetAsyncResult(IntPtr handle, object result)
        {
            var xlhandle = XLOPER12.FromObject(handle);
            var xlresult = XLOPER12.FromObject(result);
            try
            {
                using (var p1 = new StructIntPtr<XLOPER12>(ref xlhandle))
                using (var p2 = new StructIntPtr<XLOPER12>(ref xlresult))
                    AsyncReturn(p1.Ptr, p2.Ptr);
            }
            finally
            {
                xlresult.Dispose();
                xlhandle.Dispose();
            }
        }

        /// <summary>
        /// Gets/Sets the RTD throttle
        /// </summary>
        public static int RTDThrottleIntervalMilliseconds
        {
            get => App.Instance.Underlying.RTD.ThrottleInterval;
            set => App.Instance.Underlying.RTD.ThrottleInterval = value;
        }

        /// <summary>
        /// Calls the specified <see cref="IRtdServerImpl"/> server.
        /// </summary>
        /// <typeparam name="TRtdServerImpl"></typeparam>
        /// <param name="implFactory"></param>
        /// <param name="arg0"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static object RTD<TRtdServerImpl>(Func<IRtdServerImpl> implFactory
            , string arg0, params string[] args) where TRtdServerImpl : IRtdServerImpl
        {
            using (var reg = new RtdRegistry(typeof(TRtdServerImpl), implFactory))
            {
                return RTD( reg.ProgId, arg0, args );
            }
        }

        /// <summary>
        /// Calls the specified server.
        /// </summary>
        /// <param name="progId"></param>
        /// <param name="arg0"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static object RTD(string progId, string arg0, params string[] args)
        {
            var arguments = new string[] { progId, string.Empty, arg0 }
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
            unsafe
            {
                var result = (XLOPER12*)ptr.ToPointer();
                var obj = result == null ? null : result->ToObject();

                if (Environment.Is64BitProcess)
                    xlAutoFree64(ptr);
                else
                    xlAutoFree32(ptr);
                return obj;
            }
        }
    }
}
