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
using System;
using System.Collections.Generic;
using ExcelMvc.Bindings;

namespace ExcelMvc.Views
{
    /// <summary>
    /// Represents the EventArgs for a View event
    /// </summary>
    public class ViewEventArgs : EventArgs
    {
        /// <summary>
        /// Gets and sets the event specific state object
        /// </summary>
        public object State { get; set; }

        /// <summary>
        /// Gets the View associated with the event
        /// </summary>
        public View View { get; private set; }

        /// <summary>
        /// Indicates if the event has been sigalled to be cancelled by any of the event sinks
        /// </summary>
        public bool IsCancelled { get; private set; }

        /// <summary>
        /// Signals that the event should be cancelled. Note IsCancelled is not allowed to be set to false.
        /// </summary>
        public void Cancel()
        {
            IsCancelled = true;
        }

        /// <summary>
        /// Initialies an instance of  ExcelMvc.Views.ViewEventArgs
        /// </summary>
        /// <param name="view">View associated with the event</param>
        public ViewEventArgs(View view)
        {
            View = view;
        }
    }

    /// <summary>
    /// Handler for a view event
    /// </summary>
    /// <param name="sender">Event sender</param>
    /// <param name="args">Event Args</param>
    public delegate void ViewEventHandler(object sender, ViewEventArgs args);

    /// <summary>
    /// Handler for the Destroyed event
    /// </summary>
    /// <param name="sender"></param>
    public delegate void DestroyedHandler(object sender);

    /// <summary>
    /// Represents the EventArgs for a SelectionChanged event
    /// </summary>
    public class SelectionChangedArgs : EventArgs
    {
        /// <summary>
        /// Bindings selected
        /// </summary>
        public IEnumerable<Binding> Bindings { get; private set; }

        /// <summary>
        /// Items selected
        /// </summary>
        public IEnumerable<object> Items { get; private set; }

        /// <summary>
        /// Initialises an instance of SelectionChangedArgs
        /// </summary>
        /// <param name="items">Items changed</param>
        /// <param name="bindings">Bindings selected</param>
        public SelectionChangedArgs(IEnumerable<object> items, IEnumerable<Binding> bindings)
        {
            Items = items;
            Bindings = bindings;
        }
    }

    /// <summary>
    /// Handler for the SelectionChanged event
    /// </summary>
    /// <param name="sender"></param>
    public delegate void SelectionChangedHandler(object sender, SelectionChangedArgs args);


    /// <summary>
    /// Represents the EventArgs for a ObjectChanged event
    /// </summary>
    public class ObjectChangedArgs : EventArgs
    {
        /// <summary>
        /// Items changed
        /// </summary>
        public IEnumerable<object> Items { get; private set; }

        /// <summary>
        /// Properties changed
        /// </summary>
        public IEnumerable<string> Paths { get; private set; }

        /// <summary>
        /// Initialises an instance of ObjectChangedArgs
        /// </summary>
        /// <param name="items">Objects changed</param>
        /// <param name="paths">Property changed</param>
        public ObjectChangedArgs(IEnumerable<object> items, IEnumerable<string> paths)
        {
            Items = items;
            Paths = paths;
        }
    }

    /// <summary>
    /// Handler for the ObjectChanged event
    /// </summary>
    /// <param name="sender"></param>
    public delegate void ObjectChangedHandler(object sender, ObjectChangedArgs args);

}
