using ExcelMvc.Functions;
using Shared;

namespace AddInB
{
    public static class Functions
    {
        [ExcelFunction(Name = "uB")]
        public static double B(double value)
        {
            Data.Value = value;
            return value;
        }
    }
}