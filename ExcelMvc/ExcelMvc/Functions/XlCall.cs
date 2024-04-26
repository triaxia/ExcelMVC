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
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelMvc.Functions
{
    public class MessageEventArgs : EventArgs
    {
        public string Message { get; }
        public MessageEventArgs(string message)
            => Message = message;
    }

    public class RegisteringEventArgs : EventArgs
    {
        public IEnumerable<Function> Functions { get; }
        public RegisteringEventArgs(IEnumerable<Function> functions)
            => Functions = functions;
    }

    public static class XlCall
    {
        internal static void RegisterFunctions(Functions functions)
        {
            OnRegistering(functions);
            using (var pFunction = new StructIntPtr<Functions>(ref functions))
            {
                AddIn.RegisterFunctions(pFunction.Ptr);
            }
        }

        internal static void AsyncReturn(IntPtr handle, IntPtr result)
        {
            AddIn.AutoFree(AddIn.AsyncReturn(handle, result));
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
        /// Occurs before functions are registered to Excel. 
        /// </summary>
        public static event EventHandler<RegisteringEventArgs> Registering;

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
        /// Gets the Excel caller range.
        /// </summary>
        /// <returns></returns>
        public static ExcelReference GetCallerReference()
            => ExcelReference.GetCaller();

        /// <summary>
        /// Gets a reference on the workbook and worksheet specified.
        /// </summary>
        /// <param name="bookName"></param>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        /// <returns></returns>
        public static ExcelReference GetReference(string bookName, string sheetName
            , int row, int column, int rowCount, int columnCount)
            => new ExcelReference(sheetName, row, column, rowCount, columnCount);

        /// <summary>
        /// Gets a reference on the active workbook.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        /// <returns></returns>
        public static ExcelReference GetActiveBookReference(string sheetName
            , int row, int column, int rowCount, int columnCount)
            => new ExcelReference(sheetName, row, column, rowCount, columnCount);

        /// <summary>
        /// Gets a reference on the active worksheet.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="rowCount"></param>
        /// <param name="columnCount"></param>
        /// <returns></returns>
        public static ExcelReference GetActiveSheetReference(int row, int column, int rowCount, int columnCount)
            => new ExcelReference(row, column, rowCount, columnCount);

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
            get => App.Instance.Underlying?.RTD.ThrottleInterval ?? 0;
            set
            {
                if (App.Instance.Underlying != null)
                    App.Instance.Underlying.RTD.ThrottleInterval = value;
            }
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
                ptr = AddIn.RtdCall(pArgs.Ptr);
            }
            unsafe
            {
                var result = (XLOPER12*)ptr.ToPointer();
                var obj = result == null ? null : result->ToObject();
                AddIn.AutoFree(ptr);
                return obj;
            }
        }

        /// <summary>
        /// Indicates if the Excel function wizard window is open.
        /// </summary>
        /// <returns></returns>
        public static bool IsInFunctionWizard()
            => IsInteractiveEditing();

        /// <summary>
        /// Indicates if Excel is in editing mode.
        /// </summary>
        /// <returns></returns>
        public static bool IsInteractiveEditing()
        {
            var app = App.Instance.Underlying;
            if (!(app?.Interactive ?? false))
                return false;
            
            try
            {
                app.Interactive = false;
                app.Interactive = true;
            }
            catch
            {
                return true;
            }
            return false;
        }

        private static void OnRegistering(Functions functions, object sender = null)
        {
            Registering?.Invoke(sender, new RegisteringEventArgs(functions.Items));
        }
    }
}
