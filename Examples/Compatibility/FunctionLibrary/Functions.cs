using Function.Interfaces;
using Samples;

namespace FunctionLibrary
{
    public static class Functions
    {
        [Function("Add 2 numbers")]
        public static double Add([Argument("Argument x")] double x, [Argument("Argument y")] double y)
        {
            return Host.Instance.IsInFunctionWizard() ? double.MinValue : x + y;   
        }

        [Function("Add 2 numbers")]
        public static object Timer([Argument("Argument name")] string name)
        {
            return Host.Instance.IsInFunctionWizard() ? ""
                : Host.Instance.Rtd<TimerServer>(() => new TimerServer(), "", name);
        }
    }
}
