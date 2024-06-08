using ExcelDnaInterOp;
using Samples;
using System;

namespace FunctionLibrary
{
    public static class Functions
    {
        [Function("Add 2 numbers")]
        public static object Add([Argument("Argument x")] double x, [Argument("Argument y")] object y)
        {
            if (Function.Interfaces.FunctionHost.Instance.IsInFunctionWizard())
                return double.MinValue;

            if (y == Function.Interfaces.FunctionHost.Instance.ValueMissing)
                return Function.Interfaces.FunctionHost.Instance.ValueEmpty;

            return x + Convert.ToDouble(y);
        }

        [Function("Add 2 numbers")]
        public static object Timer([Argument("Argument name")] string name)
        {
            return Function.Interfaces.FunctionHost.Instance.IsInFunctionWizard() ? ""
                : Function.Interfaces.FunctionHost.Instance.Rtd<TimerServer>(() => new TimerServer(), "", name);
        }
    }
}
