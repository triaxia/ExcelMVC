/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany (2013)

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

namespace ExcelMvc.Windows
{
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows.Forms;

    /// <summary>
    /// Subclasses a window
    /// </summary>
    internal sealed class AsyncWindow : NativeWindow
    {
        private static readonly uint AsyncMessage;

        static AsyncWindow()
        {
            AsyncMessage = DllImports.RegisterWindowMessage("__ExcelMvcAsyncAction__");
        }

        /// <summary>
        /// Initializes an instance of Window
        /// </summary>
        public AsyncWindow()
        {
            var cp = new CreateParams();
            CreateHandle(cp);
        }

        /// <summary>
        /// Handler for a AsyncAction
        /// </summary>
        /// <param name="sender">Event sender</param>
        /// <param name="args">EventArgs</param>
        public delegate void AsyncMessageReceivedHandler(object sender, EventArgs args);

        /// <summary>
        /// Occurs when an async action message is received.
        /// </summary>
        public event AsyncMessageReceivedHandler AsyncMessageReceived = delegate { };

        /// <summary>
        /// Posts an async action message
        /// </summary>
        public void PostAsyncActionMessage()
        {
            var watch = Stopwatch.StartNew();
            while (watch.Elapsed.TotalSeconds < 2)
            {
                var status = DllImports.PostMessage(Handle, (int)AsyncMessage, 0, 0);
                if (status != 0) break;
                var ex = new Exception($"AsyncWindow.PostAsyncMessage failed {Marshal.GetLastWin32Error()}");
                Function.Interfaces.FunctionHost.Instance.RaiseFailed(this, new System.IO.ErrorEventArgs(ex));
                Thread.Sleep(100);
            }
        }

        /// <summary>
        /// Windows proc
        /// </summary>
        /// <param name="m">Message instance</param>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == AsyncMessage)
            {
                AsyncMessageReceived(this, EventArgs.Empty);
                return;
            }
            base.WndProc(ref m);
        }
    }
}
