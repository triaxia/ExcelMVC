using ExcelMvc.Functions;

namespace ExcelMvc.App
{
    internal class Program
    {
        static void Main(string[] args)
        {
            RunFunctionArgsMemory();
        }

        public static void RunMarshalMemory()
        {
            var context = new XlMarshalContext();
            for (int i = 0; i < 10000000; i++)
            {
                var arguments = new string[] { "ExcelMvc.Test", $"arg-{i}", $"arg-{i}{i}" };
                var text = string.Join(Environment.NewLine, arguments);
                context.ObjectToIntPtr(text);
            }
        }

        public static void RunFunctionArgsMemory()
        {
            for (int i = 0; i < 10000000; i++)
            {
                var arguments = new string[] { "ExcelMvc.Test", $"arg-{i}", $"arg-{i}{i}" }
                .Select((x, idx) => new FunctionArgument($"p{idx}", x))
                .ToArray();
                var fArgs = new FunctionArguments(0, arguments);
                using (var pArgs = new StructIntPtr<FunctionArguments>(ref fArgs))
                {
                }
            }
        }
    }
}
