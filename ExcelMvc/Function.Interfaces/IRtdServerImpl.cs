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

namespace Function.Interfaces
{
    /// <summary>
    /// Provides data for an Rtd topic.
    /// </summary>
    public class RtdTopic
    {
        /// <summary>
        /// The arguments of the topic.
        /// </summary>
        public string[] Args { get; }

        /// <summary>
        /// The value of the topic.
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Initializes a new instance of <see cref="RtdTopic"/>
        /// </summary>
        /// <param name="args"></param>
        /// <param name="value"></param>
        public RtdTopic(string[] args, object value)
        {
            Args = args;
            Value = value;
        }

        /// <inheritdoc cref="object.ToString"/>
        public override string ToString()
        {
            return $"arguments={string.Join("|", Args)}, value={Value}";
        }
    }

    /// <summary>
    /// Provides data for Rtd updated events.
    /// </summary>
    public class RtdServerUpdatedEventArgs : EventArgs
    {
        /// <summary>
        /// The server that updates the topic
        /// </summary>
        public IRtdServerImpl Impl { get; }

        /// <summary>
        /// The topics updated
        /// </summary>
        public IEnumerable<RtdTopic> Topics { get; }

        /// <summary>
        /// Initializes a new instance of <see cref="RtdServerUpdatedEventArgs"/>.
        /// </summary>
        /// <param name="impl"></param>
        /// <param name="topics"></param>
        public RtdServerUpdatedEventArgs(IRtdServerImpl impl, IEnumerable<RtdTopic> topics)
        {
            Impl = impl;
            Topics = topics;
        }
    }

    /// <summary>
    /// Defines the implementation interface of a real-time data (RTD) server.
    /// </summary>
    public interface IRtdServerImpl
    {
        /// <summary>
        /// Occurs whenever topic values have been updated.
        /// </summary>
        event EventHandler<RtdServerUpdatedEventArgs> Updated;

        /// <summary>
        /// Called immediately after a RTD server is instantiated. 
        /// </summary>
        /// <returns>A negative value or zero indicates failure to start the server, a positive value
        /// indicates success.</returns>
        int Start();

        /// <summary>
        /// Adds a new topic to the RTD server.
        /// </summary>
        /// <param name="topicId"></param>
        /// <param name="args"></param>
        /// <returns>The topic value</returns>
        object Connect(int topicId, string[] args);

        /// <summary>
        /// Notifies the RTD server that a topic is no longer in use.
        /// </summary>
        /// <param name="topicId"></param>
        void Disconnect(int topicId);

        /// <summary>
        /// Gets the updated topic values in the RTD server
        /// </summary>
        /// <returns>A array of two rows, with the first row being the topic ids and 
        /// the second being the corresponding topic values.</returns>
        object[,] GetTopicValues();

        /// <summary>
        /// Indicates if the real-time data server (RTD) is still active.
        /// </summary>
        /// <returns>Zero or a negative number indicates failure; a positive number indicates that the server is active</returns>
        int Heartbeat();

        /// <summary>
        /// Terminates the real-time data (RTD) server.
        /// </summary>
        void Stop();
    }
}
