using Microsoft.Office.Interop.Excel;
using Addin.Interfaces;
using System;
using System.Linq;

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
            void OnUpdated(object sender, EventArgs args)
            {
                CallbackObject.UpdateNotify();
            }
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
    }
}
