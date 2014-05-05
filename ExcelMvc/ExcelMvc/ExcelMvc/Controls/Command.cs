/*
Copyright (c) 2013 Peter Gu or otherwise indicated by the license information contained within
the source files.

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
using ExcelMvc.Views;

namespace ExcelMvc.Controls
{
    /// <summary>
    /// Represents EventArgs for a Command event
    /// </summary>
    public class CommandEventArgs : EventArgs
    {
        public bool Handled { get; set; }
    }

    /// <summary>
    /// Defines the handler for a clicked event
    /// </summary>
    /// <param name="sender">Objects created the Clicked event </param>
    /// <param name="args">Commard argument</param>
    public delegate void ClickedHandler(object sender, CommandEventArgs args);

    /// <summary>
    /// Defines an abstract base class for Commands
    /// </summary>
    public abstract class Command : IDisposable
    {
        /// <summary>
        /// Gets the host view 
        /// </summary>
        public View Host { get; private set; }

        /// <summary>
        /// Name of the command
        /// </summary>
        public abstract string Name { get; }

        /// <summary>
        /// Caption of the command
        /// </summary>
        public abstract string Caption { get; set; }

        /// <summary>
        /// Gets and sets the Enabled state
        /// </summary>
        public abstract bool IsEnabled { get; set; }

        /// <summary>
        /// Gets and sets the command value
        /// </summary>
        public abstract object Value { get; set; }

        /// <summary>
        /// Occurs when the command is clicked
        /// </summary>
        public event ClickedHandler Clicked = delegate { };

        /// <summary>
        /// 
        /// </summary>
        /// <param name="host"></param>
        /// <param name="underlying"></param>
        protected Command(View host)
        {
            Host = host;
        }

        /// <summary>
        /// Fires the Clicked event
        /// </summary>
        public void FireClicked()
        {
            Clicked(this, new CommandEventArgs());
        }

        public virtual void Dispose()
        {
            Host = null;
        }
    }
}
