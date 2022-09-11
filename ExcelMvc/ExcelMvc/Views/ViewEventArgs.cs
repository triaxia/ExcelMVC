
/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany

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

namespace ExcelMvc.Views
{
    using System;

    /// <summary>
    /// Handler for a view event
    /// </summary>
    /// <param name="sender">Event sender</param>
    /// <param name="args">Event Args</param>
    public delegate void ViewEventHandler(object sender, ViewEventArgs args);

    /// <summary>
    /// Represents data for a View event
    /// </summary>
    public class ViewEventArgs : EventArgs
    {
        private int acceptedCount;

        /// <summary>
        /// Initialies an instance of  ExcelMvc.Views.ViewEventArgs
        /// </summary>
        /// <param name="view">View associated with the event</param>
        public ViewEventArgs(View view)
        {
            View = view;
            acceptedCount = 0;
        }

        /// <summary>
        /// Indicates at least one event sink accepted the view
        /// </summary>
        public bool IsAccepted => acceptedCount > 0;

        /// <summary>
        /// Gets and sets the event specific state object
        /// </summary>
        public object State
        {
            get;
            set;
        }

        /// <summary>
        /// Gets and sets the event specific state object
        /// </summary>
        public object Password
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the View associated with the event
        /// </summary>
        public View View
        {
            get;
            private set;
        }

        /// <summary>
        /// Indicates the calling sink is interested in the view
        /// </summary>
        /// <param name="password">Password used to unprotected the view.</param>
        public void Accept(object password = null)
        {
            acceptedCount++;
            Password = password;
        }
    }
}
