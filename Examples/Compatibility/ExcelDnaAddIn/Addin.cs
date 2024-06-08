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
            FunctionHost.Instance = new ExcelFunctionHost
            {
                FunctionAttributeType = typeof(MvcDnaInterOp.FunctionAttribute),
                ArgumentAttributeType = typeof(MvcDnaInterOp.ArgumentAttribute)
            };
        }
    }
}
