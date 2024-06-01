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
            Host.Call.StatusBarText = "what a wonderful world it is...";
            Host.Call.Registering += (_, e) =>
            {
                // overwrite function properties...
                if (e.Function.Name == "uHelp")
                    e.Function.HelpTopic = "https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1";
            };
            XlCall.ExecutingEventRaised = true;
            Host.Call.Executing += (_, e) =>
            {
                // do fast/async usage logging here...
                XlCall.RaisePosted($"Executing {e}");
            };
            Host.Call.ExceptionToFunctionResult = e => $"{e}";
        }
    }
}
