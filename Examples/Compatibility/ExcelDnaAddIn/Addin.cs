using ExcelDna.Integration;
using ExcelMvc.Functions;
using Function.Interfaces;

namespace ExcelAddIn
{
    public class AddIn : IExcelAddIn, IFunctionAddIn
    {
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            FunctionHost.Instance = new ExcelDnaHost();
        }

        public void Close()
        {
        }

        public void Open()
        {
            FunctionHost.Instance = new ExcelFunctionHost();
            FunctionHost.Instance.FunctionAttributeType = typeof(MvcDnaInterOp.FunctionAttribute);
            FunctionHost.Instance.ArgumentAttributeType = typeof(MvcDnaInterOp.ArgumentAttribute);
        }
    }
}
