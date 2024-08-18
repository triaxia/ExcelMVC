using Function.Interfaces;

namespace Sample.AddIn
{
    public class SampleAddIn : IFunctionAddIn
    {
        public int Ranking { get; } = int.MaxValue;
        public void Close()
        {
        }

        public void Open()
        {
            FunctionHost.Instance.StatusBarText = "what a wonderful world it is...";
            FunctionHost.Instance.Registering += (_, e) =>
            {
                // overwrite function properties...
                if (e.Function.Name == "uHelp")
                    e.Function.HelpTopic = "https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1";
            };
            FunctionHost.Instance.ExecutingEventRaised = true;
            FunctionHost.Instance.Executing += (_, e) =>
            {
                // do fast/async usage logging here...
                FunctionHost.Instance.RaisePosted(this, new MessageEventArgs($"Executing {e}"));
            };
            FunctionHost.Instance.ExceptionToFunctionResult = e => $"{e}";
            FunctionHost.Instance.RtdUpdated += (_, e) =>
            {
                FunctionHost.Instance.RaisePosted(this, new MessageEventArgs($"RtdUpdated {e}"));
            };
        }
    }
}
