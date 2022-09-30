using ExcelMvc.Functions;

namespace Samples
{
    public static class Math
    {
        [ExcelFunction(Name = "uAdd", IsThreadSafe = true)]
        public static double Add(double v1, double v2, double v3)
        {
            return v1 + v2 + v3;
        }
    }
}