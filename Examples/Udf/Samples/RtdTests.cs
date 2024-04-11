using ExcelMvc.Functions;

namespace Samples
{
    public static class RtdTests
    {
        [ExcelFunction(Name = "uTimer")]
        public static object uTimer(string name)
        {
            var result = XlCall.CallRtd(typeof(TimerServer)
                , () => new TimerServer(), name);
            return result;
        }
    }
}
