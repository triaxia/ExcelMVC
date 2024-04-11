using ExcelMvc.Functions;

namespace Samples
{
    public static class RtdTests
    {
        [ExcelFunction(Name = "uTimer", IsAsync = true)]
        public static string uTimer(string name)
        {
            var result = XlCall.CallRtd(typeof(TimerServer)
                , () => new TimerServer(), name);
            return result.ToString();
        }
    }
}
