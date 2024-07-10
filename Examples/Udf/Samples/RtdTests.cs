using Function.Interfaces;

namespace Samples
{
    public static class RtdTests
    {
        static RtdTests()
        {
            FunctionHost.Instance.RtdThrottleIntervalMilliseconds = 1000;
        }
        [Function(Name = "uTimer")]
        public static double uTimer(string name)
        {
            return (double) FunctionHost.Instance.Rtd<TimerServer>(() => new TimerServer(), "", name);
        }
    }
}
