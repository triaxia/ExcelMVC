using ExcelMvc.Functions;
using Mvc;

namespace Samples
{
    public static class Math
    {
        /*
            Public Sub test()
                Dim start As Date
                start = Now()
                For i = 0 To 1000000
                    Application.Run "uAdd2", 1, 3
                Next i
    
                Dim diff As Double
                diff = (Now - start) * 24 * 60 * 60
                Debug.Print diff
            End Sub
        */
        //[ExcelFunction(Name = "uAdd3", IsThreadSafe = true, Description = "Add 3 numbers", HelpTopic = "https://www.microsoft.com")]
        //public static double Add3(
        //    [ExcelArgument(Name = "v1", Description = "argument 1")] double v1,
        //    [ExcelArgument(Name = "v2", Description = "argument 2")] double v2,
        //    [ExcelArgument(Name = "v3", Description = "argument 3")] double v3)
        //{
        //    return v1 + v2 + v3;
        //}

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

        [Function(Name = "uAdd", IsAsync = false, IsThreadSafe = false, Description = "nothing", HelpTopic = "https://www.microsoft.com")]
        public static double Add()
        {
            return 32424;
        }

        [Function(Name = "uFeed", IsAsync = false, IsThreadSafe = false, Description = "nothing", HelpTopic = "https://www.microsoft.com")]
        public static object uFeed(object value)
        {
            return value;
        }

        [Function(Name = "uTest", IsAsync = false, IsThreadSafe = false, Description = "nothing", HelpTopic = "https://www.microsoft.com")]
        public static object uTest()
        {
            return FunctionExecution.ExecuteRtd();
        }
    }
}