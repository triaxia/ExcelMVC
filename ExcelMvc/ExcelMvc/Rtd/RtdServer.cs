using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelMvc.Rtd
{
    [Guid("F80F202A-B862-4D50-AA51-F0481781CB4F")]
    [ComVisible(true)][ProgId("ExcelMvc.Rtd")]
    public class RtdServer : Microsoft.Office.Interop.Excel.IRtdServer
    {
        public IRtdServerImpl Impl { get; }

        private static IRtdServerImpl Incoming;
        public static object RtdCall(IRtdServerImpl impl, string[] args)
        {
            // system lock
            try
            {
                // only if not running
                RtdRegistration.RegisterType(typeof(RtdServer));
                RtdServer.Incoming = impl;
            }
            finally
            {

                RtdRegistration.UnregisterType(typeof(RtdServer));
            }
            return null;
        }

        public RtdServer()
        {
            Impl = Incoming;
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            return Impl.Start(() => CallbackObject.UpdateNotify());
        }

        public object ConnectData(int TopicID, ref Array Strings, ref bool GetNewValues)
        {
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

        private string GetProgId(Type type) => type.GetCustomAttributes(typeof(ProgIdAttribute), true)
            .Cast<ProgIdAttribute>().Single().Value;
    }
}
