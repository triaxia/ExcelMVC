using ExcelDna.Integration;
using Function.Interfaces;

namespace ExcelDnaAddIn
{
    public class AddIn : IExcelAddIn
    {
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            FunctionHost.Instance = new ExcelDnaHost();
        }
    }
}
