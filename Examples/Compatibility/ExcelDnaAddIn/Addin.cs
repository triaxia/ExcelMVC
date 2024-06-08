using ExcelDna.Integration;
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
            FunctionHost.Instance = new ExcelDnaInterOp.ExcelDnaHost();
        }

        public void Close()
        {
        }

        public void Open()
        {
            FunctionHost.Instance.FunctionAttributeType = typeof(ExcelDnaInterOp.FunctionAttribute);
            FunctionHost.Instance.ArgumentAttributeType = typeof(ExcelDnaInterOp.ArgumentAttribute);
        }
    }
}
