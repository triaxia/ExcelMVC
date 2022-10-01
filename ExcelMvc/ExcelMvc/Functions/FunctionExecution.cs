using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class FunctionExecution
    {
        public static Dictionary<uint, (ExcelFunction function, MethodInfo method)> Functions { get; private set; }
        public static void RegisterFunctions()
        {
            Functions = FunctionDiscovery
                .Discover()
                .ToDictionary(x => x.function.Index, x => (x.function, x.method));

            foreach (var pair in Functions)
                XlCall.Register(pair.Value.function);
        }

        public static void Execute(IntPtr args)
        {
            var fargs = Marshal.PtrToStructure<FunctionArgs>(args);
            var arguments = fargs.GetArgs();
            var method = Functions[fargs.Index].method;
            var values = method.GetParameters()
                .Select((p, idx) => Converter.ConvertIncoming(arguments[idx], p))
                .ToArray();
            var result = method.Invoke(null, values);
            Converter.ConvertOutging(result, ref fargs.Result);
        }
    }
}
