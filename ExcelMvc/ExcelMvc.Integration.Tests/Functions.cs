using Function.Interfaces;

namespace ExcelMvc.Integration.Tests
{
    public static class Functions
    {
        [Function(Name = "uBool", IsMacroType =true, Description = nameof(uBool))]
        public static bool uBool(bool v1, [Argument(Name = "[v2]")] bool ?v2 = false)
        {
            return v1 && v2.Value;
        }
    }
}