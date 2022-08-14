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
            HookEvents(true);
        }
        public void Dispose()
        {
            HookEvents(false);
        }

        private void HookEvents(bool isHook)
        {
            if (isHook)
            {
                App.Instance.Opening += Book_Opening;
                App.Instance.Opened += Book_Opened;
                App.Instance.Closing += Book_Closing;
                App.Instance.Closed += Book_Closed;
            }
            else
            {
                App.Instance.Opening -= Book_Opening;
                App.Instance.Opened -= Book_Opened;
                App.Instance.Closing -= Book_Closing;
                App.Instance.Closed -= Book_Closed;
            }
        }

        void Book_Opening(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // accept view
                args.Accept();
            }
        }

        void Book_Opened(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // assign model
                args.Accept();
                args.View.Model = new ViewModelTrading(args.View);
            }
        }

        void Book_Closing(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // allow closing
                args.Accept();
            }
        }

        void Book_Closed(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // detach model
                args.View.Model = null;
            }
        }

        private bool IsMybook(ViewEventArgs args)
        {
            return args.View.Id.CompareOrdinalIgnoreCase(BookId) == 0;
        }
    }
}
