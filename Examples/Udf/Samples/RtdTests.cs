using Function.Interfaces;

namespace Samples
{
    public static class RtdTests
    {
        static RtdTests()
        {
            Host.Instance.RTDThrottleIntervalMilliseconds = 1000;
        }
        [Function(Name = "uTimer")]
        public static object uTimer(string name)
        {
            return Host.Instance.RTD<TimerServer>(() => new TimerServer(), name);
        }
    }
}
