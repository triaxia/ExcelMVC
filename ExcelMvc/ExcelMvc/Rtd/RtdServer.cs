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
using System.IO;
using System.Linq;
using ExcelMvc.Runtime;
using Function.Interfaces;
using Microsoft.Office.Interop.Excel;

namespace ExcelMvc.Rtd
{
    public class RtdServer : IRtdServer
    {
        private IRTDUpdateEvent CallbackObject { get; set; }

        public IRtdServerImpl Impl { get; }
        public RtdServer(IRtdServerImpl impl) => Impl = impl;

        public int ServerStart(IRTDUpdateEvent callbackObject)
        {
            CallbackObject = callbackObject;
            Impl.Updated -= OnUpdated;
            Impl.Updated += OnUpdated;
            return Impl.Start();
        }

        public object ConnectData(int TopicID, ref Array Strings, ref bool GetNewValues)
        {
            GetNewValues = true;
            var args = Strings.Cast<object>().Select(x => $"{x}").ToArray();
            return Impl.Connect(TopicID, args);
        }

        public Array RefreshData(ref int TopicCount)
        {
            var values = Impl.GetTopicValues();
            TopicCount = values.GetLength(1);
            return values;
        }

        public void DisconnectData(int TopicID)
        {
            Impl.Disconnect(TopicID);
        }

        public int Heartbeat()
        {
            return Impl.Heartbeat();
        }

        public void ServerTerminate()
        {
            Impl.Stop();
            RtdRegistry.OnTerminated(this);
        }

        private void OnUpdated(object sender, RtdServerUpdatedEventArgs args)
        {
            try
            {
                Host.Instance.RaiseRtdUpdated(sender, new RtdServerUpdatedEventArgs(Impl));
            }
            catch (Exception ex)
            {
                Host.Instance.RaiseFailed(this, new ErrorEventArgs(ex));
            }

            AsyncActions.Post(state =>
            {
                try
                {
                    CallbackObject.UpdateNotify();
                }
                catch (Exception ex)
                {
                    Host.Instance.RaiseFailed(this, new ErrorEventArgs(ex));
                    OnUpdated(sender, args);
                }
            }, Impl, false);
        }
    }
}
