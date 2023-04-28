using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelMvc.Rtd
{
    public class RtdServer : Microsoft.Office.Interop.Excel.IRtdServer
    {
        private IRTDUpdateEvent CallbackObject { get; set; }

        public IRtdServerImpl Impl { get; }
        public const string ProgIdPattern = "ExcelMvc.Rtd[0-9]*";
        public string GetProgId() => this.GetType().GetCustomAttributes(typeof(ProgIdAttribute), true)
            .Cast<ProgIdAttribute>().Single().Value;

        public RtdServer()
        {
            Impl = new RtdServerImplTest();
        }

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
            Impl.Terminate();
        }
    }

    [Guid("F80F202A-B862-4D50-AA51-F0481781CB4F")][ComVisible(true)][ProgId("ExcelMvc.Rtd00")]public class RtdServer00 : RtdServer { };
}
