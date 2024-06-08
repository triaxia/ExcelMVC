using Samples;
using System;
using FunctionAttribute = MvcDnaInterOp.FunctionAttribute;
using ArgumentAttribute = MvcDnaInterOp.ArgumentAttribute;
using Function.Interfaces;

namespace FunctionLibrary
{
    public static class Functions
    {
        [Function("Add 2 numbers")]
        public static object Add([Argument("Argument x")] double x, [Argument("Argument y")] object y)
        {
            if (FunctionHost.Instance.IsInFunctionWizard())
                return double.MinValue;

            if (y == FunctionHost.Instance.ValueMissing)
                return FunctionHost.Instance.ValueEmpty;

            return x + Convert.ToDouble(y);
        }

        [Function("Add 2 numbers")]
        public static object Timer([Argument("Argument name")] string name)
        {
            return FunctionHost.Instance.IsInFunctionWizard() ? ""
                : FunctionHost.Instance.Rtd<TimerServer>(() => new TimerServer(), "", name);
        }
    }
}
