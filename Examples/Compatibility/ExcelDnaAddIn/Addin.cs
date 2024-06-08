using ExcelDna.Integration;
using Function.Interfaces;
using FunctionLibrary;

namespace ExcelDnaAddIn
{
    public class Addin : IExcelAddIn
    {
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            Host.Instance = new ExcelDnaHost();
        }
    }
}
