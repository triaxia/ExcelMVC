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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Function.Interfaces
{
    /// <summary>
    /// Provides data for message events.
    /// </summary>
    public class MessageEventArgs : EventArgs
    {
        /// <summary>
        /// The message text.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Initialises a new instance of <see cref="MessageEventArgs"/>.
        /// </summary>
        /// <param name="message"></param>
        public MessageEventArgs(string message)
            => Message = message;

        /// <inheritdoc cref="object.ToString"/>
        public override string ToString()
        {
            return $"Message:{Message}";
        }
    }

    /// <summary>
    /// Provides data for function registration events.
    /// </summary>
    public class RegisteringEventArgs : EventArgs
    {
        /// <summary>
        /// The function being registered.
        /// </summary>
        public FunctionDefinition Function;

        /// <summary>
        /// Initialises a new instance of <see cref="RegisteringEventArgs"/>.
        /// </summary>
        /// <param name="function"></param>
        public RegisteringEventArgs(FunctionDefinition function)
            => Function = function;

        /// <inheritdoc cref="object.ToString"/>
        public override string ToString()
        {
            return $"Function:{Function.Name}";
        }
    }

    /// <summary>
    /// Provides data for function executing events.
    /// </summary>
    public class ExecutingEventArgs : EventArgs
    {
        /// <summary>
        /// The name of the function being executed.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// The arguments being passed to the function.
        /// </summary>
        public (string Name, object Value)[] Args { get; }

        /// <summary>
        /// Initialises a new instance of <see cref="ExecutingEventArgs"/>.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="method"></param>
        /// <param name="args"></param>
        public ExecutingEventArgs(string name, MethodInfo method, object[] args)
        {
            Name = name;
            Args = method.GetParameters()
                .Select((p, i) => (name: p.Name, value: args[i]))
                .ToArray();
        }

        /// <inheritdoc cref="object.ToString"/>
        public override string ToString()
        {
            var args = string.Join(",", Args.Select(x => $"{x.Name}={x.Value}"));
            return $"{Name}[{args}]";
        }
    }

    /// <summary>
    /// Provides functionality for calling functions on their hosts.
    /// </summary>
    public interface IFunctionHost
    {
        /// <summary>
        /// Gets/Sets the host application object.
        /// </summary>
        object Application { get; set; }

        /// <summary>
        /// Gets the object that represents a missing value.
        /// </summary>
        object ValueMissing { get; }

        /// <summary>
        /// Gets the object that represents an empty value.
        /// </summary>
        object ValueEmpty { get; }

        /// <summary>
        /// Gets the object that represents a null error.
        /// </summary>
        object ErrorNull { get; }

        /// <summary>
        /// Gets the object that represents a div0 error.
        /// </summary>
        object ErrorDiv0 { get; }

        /// <summary>
        /// Gets the object that represents a value error.
        /// </summary>
        object ErrorValue { get; }

        /// <summary>
        /// Gets the object that represents a ref error.
        /// </summary>
        object ErrorRef { get; }

        /// <summary>
        /// Gets the object that represents a name error.
        /// </summary>
        object ErrorName { get; }

        /// <summary>
        /// Gets the object that represents a num error.
        /// </summary>
        object ErrorNum { get; }

        /// <summary>
        /// Gets the object that represents a N/A error.
        /// </summary>
        object ErrorNA { get; }

        /// <summary>
        /// Gets the object that represents a data error.
        /// </summary>
        object ErrorData { get; }

        /// <summary>
        /// <see cref="ErrorValue"/>, <see cref="ErrorNull"/>... etc to their string representations./>
        /// </summary>
        IDictionary<object, string> ErrorMappings { get; }

        /// <summary>
        /// Converts error objects to their string representations.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        string ErrorToString(object value);

        /// <summary>
        /// Occurs whenever Rtd servers are updated.
        /// </summary>
        event EventHandler<RtdServerUpdatedEventArgs> RtdUpdated;

        /// <summary>
        /// Raises a <see cref="RtdUpdated"/> event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        void RaiseRtdUpdated(object sender, RtdServerUpdatedEventArgs args);

        /// <summary>
        /// Gets/Sets the RTD throttle.
        /// </summary>
        int RtdThrottleIntervalMilliseconds { get; set; }

        /// <summary>
        /// Calls the specified <see cref="ITDServerImpl"/> server.
        /// </summary>
        /// <typeparam name="TRtdServerImpl"></typeparam>
        /// <param name="implFactory"></param>
        /// <param name="server"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        object Rtd<TRtdServerImpl>(Func<IRtdServerImpl> implFactory
            , string server, params string[] args) where TRtdServerImpl : IRtdServerImpl;

        /// <summary>
        /// Calls the specified server.
        /// </summary>
        /// <param name="progId"></param>
        /// <param name="server"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        object Rtd(string progId, string server, params string[] args);

        /// <summary>
        /// Indicates if the function wizard window is open.
        /// </summary>
        /// <returns></returns>
        bool IsInFunctionWizard();

        /// <summary>
        /// Gets/Sets the host status bar text.
        /// </summary>
        string StatusBarText { get; set; }

        /// <summary>
        /// Gets the asynchronous handle.
        /// </summary>
        /// <param name="handle"></param>
        /// <returns></returns>
        IntPtr GetAsyncHandle(IntPtr handle);

        /// <summary>
        /// Sets the asynchronous result.
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="value"></param>
        void SetAsyncValue(IntPtr handle, object value);

        /// <summary>
        /// Occurs whenever messages are posted.
        /// </summary>
        event EventHandler<MessageEventArgs> Posted;

        /// <summary>
        /// Raises a <see cref="Posted"/> event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        void RaisePosted(object sender, MessageEventArgs args);

        /// <summary>
        /// Occurs before functions are registered with the host. 
        /// </summary>
        event EventHandler<RegisteringEventArgs> Registering;

        /// <summary>
        /// Raises a <see cref="Registering"/> event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        void RaiseRegistering(object sender, RegisteringEventArgs args);

        /// <summary>
        /// Occurs whenever errors are encountered.
        /// </summary>
        event EventHandler<ErrorEventArgs> Failed;

        /// <summary>
        /// Raises a <see cref="Failed"/> event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        void RaiseFailed(object sender, ErrorEventArgs args);

        /// <summary>
        /// Gets/Sets the flag indicating whether the <see cref="Executing"/> event is
        /// raised or not.
        /// </summary>
        bool ExecutingEventRaised { get; set; }

        /// <summary>
        /// Occurs whenever functions are executed.
        /// </summary>
        event EventHandler<ExecutingEventArgs> Executing;

        /// <summary>
        /// Raise an <see cref="Executing"/> event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        void RaiseExecuting(object sender, ExecutingEventArgs args);

        /// <summary>
        /// Gets/Sets the function that converts an exception to object.
        /// </summary>
        Func<Exception, object> ExceptionToFunctionResult { get; set; }

        /// <summary>
        /// Gets the range of the caller.
        /// </summary>
        /// <returns></returns>
        RangeReference GetCallerReference();

        /// <summary>
        /// Gets a reference on the book and page specified.
        /// </summary>
        /// <param name="bookName"></param>
        /// <param name="pageName"></param>
        /// <param name="rowFirst"></param>
        /// <param name="rowLast"></param>
        /// <param name="columnFirst"></param>
        /// <param name="columnLast"></param>
        /// <returns></returns>
        RangeReference GetReference(string bookName, string pageName
            , int rowFirst, int rowLast, int columnFirst, int columnLast);

        /// <summary>
        /// Gets a reference on the active book.
        /// </summary>
        /// <param name="pageName"></param>
        /// <param name="rowFirst"></param>
        /// <param name="rowLast"></param>
        /// <param name="columnFirst"></param>
        /// <param name="columnLast"></param>
        /// <returns></returns>
        RangeReference GetActiveBookReference(string pageName
            , int rowFirst, int rowLast, int columnFirst, int columnLast);

        /// <summary>
        /// Gets a reference on the active sheet.
        /// </summary>
        /// <param name="rowFirst"></param>
        /// <param name="rowLast"></param>
        /// <param name="columnFirst"></param>
        /// <param name="columnLast"></param>
        /// <returns></returns>
        RangeReference GetActiveSheetReference(int rowFirst, int rowLast, int columnFirst, int columnLast);

        /// <summary>
        /// Gets the value of the specified range.
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        object GetRangeValue(RangeReference range);

        /// <summary>
        /// Sets the value of the specified range.
        /// </summary>
        /// <param name="range"></param>
        /// <param name="value"></param>
        /// <param name="async"></param>
        void SetRangeValue(RangeReference range, object value, bool async);

        /// <summary>
        /// Registers the specified functions with the host.
        /// </summary>
        /// <param name="functions"></param>
        void RegisterFunctions(FunctionDefinitions functions);

        /// <summary>
        /// Gets/Sets the type of <see cref="FunctionAttribute"/>
        /// </summary>
        Type FunctionAttributeType { get; set; }

        /// <summary>
        /// Gets/Sets the type of <see cref="ArgumentAttribute"/>
        /// </summary>
        Type ArgumentAttributeType { get; set; }

        /// <summary>
        /// Posts an async action to the main thread of the <see cref="Underlying"/> host.
        /// </summary>
        /// <param name="action"></param>
        /// <param name="state"></param>
        void Post(Action<object> action, object state);

        /// <summary>
        /// Gets the version of the <see cref="Underlying"/> host.
        /// </summary>
        string Version { get; }

        /// <summary>
        /// Gets the flag indicating if the host IDE is open.
        /// </summary>
        bool IsIdeOpen { get; }

        /// <summary>
        /// Gets the full file name of the moduel that runs the host.
        /// </summary>
        string GetModuleFileName { get; }
    }

    /// <summary>
    /// </summary>
    public static class FunctionHost
    {
        /// <summary>
        /// Gets/Sets the implementation of <see cref="IFunctionHost"/>.
        /// </summary>
        public static IFunctionHost Instance { get; set; }
    }
}

