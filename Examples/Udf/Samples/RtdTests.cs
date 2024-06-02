using Function.Interfaces;

namespace Samples
{
    public static class RtdTests
    {
        static RtdTests()
        {
            Host.Instance.RtdThrottleIntervalMilliseconds = 1000;
        }
        [Function(Name = "uTimer")]
        public static object uTimer(string name)
        {
            return Host.Instance.Rtd<TimerServer>(() => new TimerServer(), "", name);
        }
    }
}
