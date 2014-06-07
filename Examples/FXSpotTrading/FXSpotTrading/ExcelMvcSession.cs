using System.Diagnostics;
using ExcelMvc.Runtime;
using ExcelMvc.Views;

namespace FXSpotTrading
{
    public class ExcelMvcSession : ISession
    {
        public ExcelMvcSession()
        {
            App.Instance.Opening += Instance_Opening;
            App.Instance.Opened += Instance_Opened;
            App.Instance.Closing += Instance_Closing;
            App.Instance.Closed += Instance_Closed;
        }

        void Instance_Opening(object sender, ViewEventArgs args)
        {
            if (args.View.Id != "FXSpotTrading")
                args.Cancel();
        }

        void Instance_Opened(object sender, ViewEventArgs args)
        {
        }

        void Instance_Closing(object sender, ViewEventArgs args)
        {
        }
        void Instance_Closed(object sender, ViewEventArgs args)
        {
        }

        public void Dispose()
        {
        }
    }
}
