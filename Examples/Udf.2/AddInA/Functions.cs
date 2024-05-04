using ExcelMvc.Functions;
using Shared;

namespace AddInA
{
    public static class Functions
    {

        [ExcelFunction(Name = "uA")]
        public static double A(double value)
        {
            Data.Value = value;
            return value;
        }
    }
}
