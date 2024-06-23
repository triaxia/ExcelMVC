using ExcelMvc.Functions;
using Function.Interfaces;

namespace ExcelAddIn
{
    public class AddIn : IFunctionAddIn
    {
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
