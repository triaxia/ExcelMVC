using System.Collections.Generic;
using ExcelMvc.Runtime;
using ExcelMvc.Views;

namespace FXSpotTrading.ViewModels
{
    public class ViewModelSession : ISession
    {
        private readonly Dictionary<View, object> sessions; 

        public ViewModelSession()
        {
            // hook book view notificaton event
            App.Instance.Opening += Instance_Opening;
            App.Instance.Opened += Instance_Opened;
            App.Instance.Closing += Instance_Closing;
            App.Instance.Closed += Instance_Closed;
            sessions = new Dictionary<View, object>();
        }

        void Instance_Opening(object sender, ViewEventArgs args)
        {
            // cancel out for non-ExcelMvc books
            if (args.View.Id != "FXSpotTrading")
                args.Cancel();
        }

        void Instance_Opened(object sender, ViewEventArgs args)
        {
            // create book model
            if (args.View.Id == "FXSpotTrading")
                sessions[args.View] = new ViewModelTrading(args.View);
        }

        void Instance_Closing(object sender, ViewEventArgs args)
        {
            // cancel out here
            // args.Cancel();
        }

        void Instance_Closed(object sender, ViewEventArgs args)
        {
            // remove view models
            sessions.Remove(args.View);
        }

        public void Dispose()
        {
            sessions.Clear();
        }
    }
}
