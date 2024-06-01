using Function.Interfaces;

namespace Samples
{
    public static class RtdTests
    {
        static RtdTests()
        {
            Host.Call.RTDThrottleIntervalMilliseconds = 1000;
        }
        [Function(Name = "uTimer")]
        public static object uTimer(string name)
        {
            var result = Host.Call.RTD<TimerServer>(() => new TimerServer(), name);
            return "";
        }
    }
}
