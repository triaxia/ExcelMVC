using ExcelMvc.Functions;
using Function.Interfaces;

namespace Sample.AddIn
{
    public class SampleAddIn : IAddIn
    {
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            Host.Instance.StatusBarText = "what a wonderful world it is...";
            Host.Instance.Registering += (_, e) =>
            {
                // overwrite function properties...
                if (e.Function.Name == "uHelp")
                    e.Function.HelpTopic = "https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1";
            };
            Host.Instance.ExecutingEventRaised = true;
            Host.Instance.Executing += (_, e) =>
            {
                // do fast/async usage logging here...
                XlCall.RaisePosted($"Executing {e}");
            };
            Host.Instance.ExceptionToFunctionResult = e => $"{e}";
        }
    }
}
