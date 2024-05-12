using ExcelMvc.Functions;

namespace Sample.AddIn
{
    public class SampleAddIn : IExcelAddIn
    {
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            XlCall.Registering += (_, e) =>
            {
                // overwrite function properties...
                if (e.Function.Name == "uHelp")
                    e.Function.HelpTopic = "https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1";
            };

            XlCall.Executing += (_, e) =>
            {
                // do fast/async usage logging here...
                XlCall.RaisePosted($"Executing {e.Name}");
            };
        }
    }
}
