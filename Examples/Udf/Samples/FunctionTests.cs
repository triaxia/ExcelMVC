﻿using ExcelMvc.Functions;
using Addin.Interfaces;

namespace Samples
{
    public static class FunctionTests
    {
        [Function(Name = "uAdd2", IsAsync = true, IsThreadSafe = false, Description = "Add 2 numbers", HelpTopic = "https://www.microsoft.com")]
        public static double Add2(
            [Argument(Name = "v1", Description = "argument 1")] double v1,
            [Argument(Name = "v2", Description = "argument 2")] double v2)
        {
            return v1 + v2;
        }

        [Function(Name = "uAdd3", IsAsync = false, IsThreadSafe = false, Description = "Add 2 numbers", HelpTopic = "https://www.microsoft.com")]
        public static double Add3(
            [Argument(Name = "v1", Description = "argument 1")] double v1,
            [Argument(Name = "v2", Description = "argument 2")] double v2,
            [Argument(Name = "v3", Description = "argument 3")] double v3)
        {
            return v1 + v2 + v3;
        }

        [Function(Name = "uArg", IsAsync = false, IsThreadSafe = false, Description = "Returns Args", HelpTopic = "https://www.microsoft.com")]
        public static object uFeed(object value)
        {
            return value;
        }

        [Function(Name = "uRtdTime", IsAsync = false, IsThreadSafe = false, Description = "Timer Rtd", HelpTopic = "https://www.microsoft.com")]
        public static object uRtdTime(
            [Argument(Name = "v1", Description = "argument 1")] string v1,
            [Argument(Name = "v2", Description = "argument 2")] string v2
)
        {
            return FunctionExecution.ExecuteRtd(typeof(TimerServer), v1, v2);
        }
        [Function(Name = "uRtdRandom", IsAsync = false, IsThreadSafe = false, Description = "Ramdon Rtd", HelpTopic = "https://www.microsoft.com")]
        public static object uRtdRandom(
            [Argument(Name = "v1", Description = "argument 1")] string v1,
            [Argument(Name = "v2", Description = "argument 2")] string v2
)
        {
            return FunctionExecution.ExecuteRtd(typeof(RandomTimer), v1, v2);
        }
    }
}