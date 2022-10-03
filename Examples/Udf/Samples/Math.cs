using ExcelMvc.Functions;

namespace Samples
{
    public static class Math
    {
        //[ExcelFunction(Name = "uAdd3", IsThreadSafe = true, Description = "Add 3 numbers", HelpTopic = "https://www.microsoft.com")]
        //public static double Add3(
        //    [ExcelArgument(Name = "v1", Description = "argument 1")] double v1,
        //    [ExcelArgument(Name = "v2", Description = "argument 2")] double v2,
        //    [ExcelArgument(Name = "v3", Description = "argument 3")] double v3)
        //{
        //    return v1 + v2 + v3;
        //}

        [ExcelFunction(Name = "uAdd2", IsAnyc = true, IsThreadSafe = false, Description = "Add 2 numbers", HelpTopic = "https://www.microsoft.com")]
        public static double Add2(
            [ExcelArgument(Name = "v1", Description = "argument 1")] double v1,
            [ExcelArgument(Name = "v2", Description = "argument 2")] double v2)
        {
            return v1 + v2;
        }
    }
}