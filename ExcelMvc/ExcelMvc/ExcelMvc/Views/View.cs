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

    using ExcelMvc.Bindings;
    using ExcelMvc.Controls;
    using ExcelMvc.Extensions;

    using Action = System.Action;

    /// <summary>
    /// Represents the base behaviour of Views
    /// </summary>
    public abstract class View : IDisposable
    {
        #region Events

        /// <summary>
        /// Occurs after a View is activated. 
        /// </summary>
        public event ViewEventHandler Activated = delegate { };

        /// <summary>
        /// Occurs when a binding exception is caught
        /// </summary>
        public event BindingFailedHandler BindFailed = delegate { };

        /// <summary>
        /// Occurs after a View is closed. 
        /// </summary>
        public event ViewEventHandler Closed = delegate { };

        /// <summary>
        ///  Occurs before a View is closed. 
        /// </summary>
        public event ViewEventHandler Closing = delegate { };

        /// <summary>
        /// Occurs after a View is activated. 
        /// </summary>
        public event ViewEventHandler Deactivated = delegate { };

        /// <summary>
        /// Occurs when a View has been destroyed
        /// </summary>
        public event DestroyedHandler Destroyed = delegate { };

        /// <summary>
        /// Occurs when a view's objects are changed
        /// </summary>
        public event ObjectChangedHandler ObjectChanged = delegate { };

        /// <summary>
        /// Occurs after a View is opened. 
        /// </summary>
        public event ViewEventHandler Opened = delegate { };

        /// <summary>
        /// Occurs before a View is opened. 
        /// </summary>
        public event ViewEventHandler Opening = delegate { };

        /// <summary>
        /// Occurs when a view's selection is changed
        /// </summary>
        public event SelectionChangedHandler SelectionChanged = delegate { };

        #endregion Events

        #region Properties

        /// <summary>
        /// Gets the child views
        /// </summary>
        public virtual IEnumerable<View> Children
        {
            get
            {
                return new View[] { };
            }
        }

        /// <summary>
        /// Gets the commands
        /// </summary>
        public virtual IEnumerable<Command> Commands
        {
            get
            {
                return new Command[] { };
            }
        }

        /// <summary>
        /// Gets the view id
        /// </summary>
        public abstract string Id
        {
            get;
        }

        /// <summary>
        /// Gets and sets the underlying model
        /// </summary>
        public virtual object Model
        {
            get; set;
        }

        /// <summary>
        /// Gets the view name
        /// </summary>
        public abstract string Name
        {
            get;
        }

        /// <summary>
        /// Gets the parent view
        /// </summary>
        public View Parent
        {
            get; protected set;
        }

        /// <summary>
        /// Gets the view type
        /// </summary>
        public abstract Binding.ViewType Type
        {
            get;
        }

        #endregion Properties

        #region Methods

        public abstract void Dispose();

        /// <summary>
        /// Finds the view with the name specified, starting from this instance and downwards
        /// </summary>
        /// <param name="name">Name of the view</param>
        /// <returns>View found or null</returns>
        public View Find(string name)
        {
            return Find(Binding.ViewType.None, name);
        }

        /// <summary>
        /// Finds the view with the name specified, starting from this instance and downwards
        /// </summary>
        /// <param name="type">View type</param>
        /// <param name="name">Name of the view</param>
        /// <returns>View found or null</returns>
        public View Find(Binding.ViewType type, string name)
        {
            if ((type == Binding.ViewType.None || type == Type)
                && Name.CompareOrdinalIgnoreCase(name) == 0)
                return this;

            View result = null;
            foreach (var child in Children)
            {
                result = child.Find(type, name);
                if (result != null)
                    break;
            }

            return result;
        }

        /// <summary>
        /// Finds a command
        /// </summary>
        /// <param name="name">Name of the command to find</param>
        /// <returns>Command found or null</returns>
        public Command FindCommand(string name)
        {
            foreach (var cmd in Commands)
            {
                if (cmd.Name.CompareOrdinalIgnoreCase(name) == 0)
                    return cmd;
            }

            foreach (var child in Children)
            {
                var cmd = child.FindCommand(name);
                if (cmd != null)
                    return cmd;
            }

            return null;
        }

        /// <summary>
        /// Hooks a binding failed handler
        /// </summary>
        /// <param name="handler">Handler to be hooked</param>
        /// <param name="isHook">Indicates if this call is to hook or unhook the handler</param>
        public void HookBindingFailed(BindingFailedHandler handler, bool isHook)
        {
            if (isHook)
                BindFailed += handler;
            else
                BindFailed -= handler;

            foreach (var child in Children)
                child.HookBindingFailed(handler, isHook);
        }

        /// <summary>
        /// Hooks a clicked handler to commands
        /// </summary>
        /// <param name="handler">Handled to be hooked</param>
        /// <param name="isHook">Indicates if this call is to hook or unhook the handler</param>
        public void HookClicked(ClickedHandler handler, bool isHook)
        {
            HookClicked(handler, null, isHook);
        }

        /// <summary>
        /// Hooks a clicked handler to commands
        /// </summary>
        /// <param name="handler">Handled to be hooked</param>
        /// <param name="name">Command name</param>
        /// <param name="isHook">Indicates if this call is to hook or unhook the handler</param>
        public void HookClicked(ClickedHandler handler, string name, bool isHook)
        {
            foreach (var cmd in Commands)
            {
                if (name == null || cmd.Name.CompareOrdinalIgnoreCase(name) == 0)
                {
                    if (isHook)
                        cmd.Clicked += handler;
                    else
                        cmd.Clicked -= handler;
                }
            }

            foreach (var child in Children)
                child.HookClicked(handler, name, isHook);
        }

        /// <summary>
        /// Fires the Activated event
        /// </summary>
        /// <param name="args">Event args</param>
        public void OnActivated(ViewEventArgs args)
        {
            Activated(this, args);
        }

        /// <summary>
        /// Fires the Closed event
        /// </summary>
        /// <param name="args">Event args</param>
        public void OnClosed(ViewEventArgs args)
        {
            Closed(this, args);
        }

        /// <summary>
        /// Fires the Closing event
        /// </summary>
        /// <param name="args">Event args</param>
        public void OnClosing(ViewEventArgs args)
        {
            Closing(this, args);
        }

        /// <summary>
        /// Fires the Deactivated event
        /// </summary>
        /// <param name="args">Event args</param>
        public void OnDeactivated(ViewEventArgs args)
        {
            Deactivated(this, args);
        }

        /// <summary>
        /// Fires the Destroyed event
        /// </summary>
        /// <param name="sender">Sender</param>
        public void OnDestroyed(object sender)
        {
            Destroyed(sender);
        }

        /// <summary>
        /// Fires the ObjectChanged event
        /// </summary>
        /// <param name="items">Items changed</param>
        /// <param name="paths">Paths changed</param>
        public void OnObjectChanged(IEnumerable<object> items, IEnumerable<string> paths)
        {
            ObjectChanged(this, new ObjectChangedArgs(items, paths));
        }

        /// <summary>
        /// Fires the Opened event
        /// </summary>
        /// <param name="args">Event args</param>
        public void OnOpened(ViewEventArgs args)
        {
            Opened(this, args);
        }

        /// <summary>
        /// Fires the Opening event
        /// </summary>
        /// <param name="args">Event args</param>
        public void OnOpening(ViewEventArgs args)
        {
            Opening(this, args);
        }

        /// <summary>
        /// Fires the SelectionChanged event
        /// </summary>
        /// <param name="items">Items selected</param>
        /// <param name="bindings">Bindings selected</param>
        public void OnSelectionChanged(IEnumerable<object> items, IEnumerable<Binding> bindings)
        {
            SelectionChanged(this, new SelectionChangedArgs(items, bindings));
        }

        /// <summary>
        /// Executes an binding exception and raise BindFailed event if an exception is caught
        /// </summary>
        /// <param name="action">Action to be executed</param>
        protected void ExcuteBinding(Action action)
        {
            ActionExtensions.Try(action, ex => BindFailed(this, new BindingFailedEventArgs(this, ex)));
        }

        #endregion Methods
    }
}