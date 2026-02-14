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

namespace ExcelMvc.Runtime
{
    using System;
    using System.Collections.Generic;
    using ExcelMvc.Diagnostics;
    using ExcelMvc.Windows;
    using Function.Interfaces;

    /// <summary>
    /// Posts and handles asynchronous actions.
    /// </summary>
    internal static class AsyncActions
    {
        private class Item
        {
            public Action<object> Action { get; set; }
            public object State { get; set; }
        }
        private static AsyncWindow Context { get; set; }
        private static Queue<Item> Actions { get; set; }

        static AsyncActions()
        {
            Context = new AsyncWindow();
            Actions = new Queue<Item>();
            Context.AsyncMessageReceived += MainWindow_AsyncMessageReceived;
        }

        static void MainWindow_AsyncMessageReceived(object sender, EventArgs args)
        {
            Execute();
        }

        /// <summary>
        /// Initialise class static states
        /// </summary>
        public static void Initialise()
        {
            MessageWindow.CreateInstance();
        }

        /// <summary>
        /// Posts an Async macro
        /// </summary>
        /// <param name="action">Action to be executed</param>
        /// <param name="state">State object</param>
        public static void Post(Action<object> action, object state)
        {
            lock (Actions)
            {
                var wasEmpty = Actions.Count == 0;
                var item = new Item { Action = action, State = state };
                Actions.Enqueue(item);
                if (wasEmpty) Context.PostAsyncActionMessage();
            }
        }

        /// <summary>
        /// Executes the next action in the queue
        /// </summary>
        internal static void Execute()
        {
            lock (Actions)
            {
                while (Actions.Count > 0)
                {
                    Item item = Actions.Dequeue();
                    try
                    {
                        item.Action(item.State);
                    }
                    catch (Exception ex)
                    {
                        FunctionHost.Instance.RaiseFailed(item, new System.IO.ErrorEventArgs(ex));
                    }
                }
            }
        }
    }
}
