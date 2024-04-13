using ExcelMvc.Functions;

namespace Samples
{
    public static class RtdTests
    {
        static RtdTests()
        {
            XlCall.RTDThrottleIntervalMilliseconds = 1000;
        }
        [ExcelFunction(Name = "uTimer")]
        public static object uTimer(string name)
        {
            var result = XlCall.RTD(typeof(TimerServer)
                , () => new TimerServer(), name);
            return result;
        }
    }
}
