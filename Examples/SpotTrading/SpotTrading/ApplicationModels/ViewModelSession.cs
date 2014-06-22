namespace SpotTrading.ApplicationModels
{
    using ExcelMvc.Extensions;
    using ExcelMvc.Runtime;
    using ExcelMvc.Views;

    public class ViewModelSession : ISession
    {
        private const string BookId = "SpotTrading";
        public ViewModelSession()
        {
            // hook notificaton events
            App.Instance.Opening += Instance_Opening;
            App.Instance.Opened += Instance_Opened;
            App.Instance.Closing += Instance_Closing;
            App.Instance.Closed += Instance_Closed;
        }

        void Instance_Opening(object sender, ViewEventArgs args)
        {
            // cancel out for non-ExcelMvc books
            if (args.View.Id.CompareOrdinalIgnoreCase(BookId) != 0)
                args.Cancel();
        }

        void Instance_Opened(object sender, ViewEventArgs args)
        {
            // create book model
            if (args.View.Id.CompareOrdinalIgnoreCase(BookId) == 0)
                args.View.Model = new ViewModelTrading(args.View);
        }

        void Instance_Closing(object sender, ViewEventArgs args)
        {
            // cancel close
            // args.Cancel();
        }

        void Instance_Closed(object sender, ViewEventArgs args)
        {
            // remove view models
            args.View.Model = null;
        }

        public void Dispose()
        {
        }
    }
}
