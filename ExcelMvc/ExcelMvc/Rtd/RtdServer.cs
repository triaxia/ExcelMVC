using Microsoft.Office.Interop.Excel;
using ExcelMvc.Interfaces;
using System;
using System.Linq;
using ExcelMvc.Functions;
using System.Reflection;
using System.Threading.Tasks;

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

        public static object ExecuteRtd(Type implType, Func<IRtdServerImpl> implFactory, params string[] args)
        {
            using (var reg = new RtdRegistry(implType, implFactory))
            {
                args = new string[] { reg.ProgId, "" }.Concat(args).ToArray();
                var x = new FunctionArgsBag(args);
                {
                    var fargs = x.ToArgs();
                    using (var p = new StructIntPtr<FunctionArgs>(ref fargs))
                    {

                        var result = XLOPER12.FromIntPtr(XlCall.RtdCall(p.Ptr));
                        return result == null ? null : XLOPER12.ToObject(result.Value);
                    }
                }
            }
        }

        public static void SetAsyncResult(IntPtr handle, object result)
        {
            var outcome = XLOPER12.FromObject(result);
            using (var ptr = new StructIntPtr<XLOPER12>(ref outcome))
                XlCall.AsyncReturn(handle, ptr.Detach());
        }
    }
}
