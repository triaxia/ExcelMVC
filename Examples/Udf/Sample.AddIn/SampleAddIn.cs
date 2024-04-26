using ExcelMvc.Functions;
using System.Linq;

namespace Sample.AddIn
{
    internal class SampleAddIn : IExcelAddIn
    {
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            XlCall.Registering += (_, e) =>
            {
                /*
                var x = e.Functions.ToArray();
                foreach (var function in x )
                {
                    var d = e.Functions.Last();
                    d.HelpTopic = "e.Functions.ToArray(";
                }
                */
            };
        }
    }
}
