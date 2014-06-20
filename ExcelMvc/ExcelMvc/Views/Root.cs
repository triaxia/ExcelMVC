#region Header
/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Developer:         Wolfgang Stamm, Germany

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
#endregion Header

namespace ExcelMvc.Views
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    /// <summary>
    /// Wraps an native window
    /// </summary>
    public class Root : NativeWindow
    {
        #region Fields

        private static readonly uint AsyncUpdateMsg;
        private static readonly Dictionary<int, Action<object>> Actions = new Dictionary<int, Action<object>>();
        private static readonly Dictionary<int, object> States = new Dictionary<int, object>();
    
        #endregion Fields

        #region Constructors

        static Root()
        {
            AsyncUpdateMsg = RegisterWindowMessage("__ExcelMvcAsyncUpdate__");
        }

        /// <summary>
        /// Intialises an instance of Window
        /// </summary>
        /// <param name="hwnd">Window handle</param>
        public Root(int hwnd)
        {
            AssignHandle(new IntPtr(hwnd));
        }

        #endregion Constructors

        #region Delegates

        /// <summary>
        /// Handler for a Destroyed event
        /// </summary>
        /// <param name="sender">Event sender</param>
        /// <param name="args">EventArgs</param>
        public delegate void DestroyedHandler(object sender, EventArgs args);

        #endregion Delegates

        #region Events

        /// <summary>
        /// Occurs when a Window has been destroyed
        /// </summary>
        public event DestroyedHandler Destroyed = delegate { };

        #endregion Events

        #region Methods


        /// <summary>
        /// Performs an Asnc action
        /// </summary>
        /// <param name="action">Action to be executed</param>
        /// <param name="state">State object</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public void Post(Action<object> action, object state)
        {
            var key = Actions.Count;
            Actions[key] = action;
            States[key] = state;
            PostMessage(Handle, (int)AsyncUpdateMsg, key, 0);
        }

        protected override void WndProc(ref Message m)
        {
            const int wmDestroy = 0x0002;
            if (m.Msg == wmDestroy)
            {
                Destroyed(this, new EventArgs());
            }
            else if (m.Msg == AsyncUpdateMsg)
            {
                Act((int)m.WParam);
                return;
            }

            base.WndProc(ref m);
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        private void Act(int key)
        {
            try
            {
                Actions[key](States[key]);
            }
            finally
            {
                Actions.Remove(key);
                States.Remove(key);
            }
        }

        [DllImport("user32.dll")]
        private static extern int PostMessage(IntPtr hwnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterWindowMessage(string lpProcName);

        #endregion Methods
    }
}