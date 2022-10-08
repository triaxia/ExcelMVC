using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

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
                XlCall.RegisterFunction(pair.Value.function);
        }

        public static void Execute(IntPtr args)
        {
            var fargs = Marshal.PtrToStructure<FunctionArgs>(args);
            var (function, method) = Functions[fargs.Index];
            var arguments = fargs.GetArgs();
            var values = method.GetParameters()
                .Select((p, idx) => Converter.ConvertIncoming(arguments[idx], p))
                .ToArray();

            if (function.IsAnyc)
                ExecuteAsync(function, method, values, ref fargs.Result);
            else
                ExecuteSync(function, method, values, ref fargs.Result);
        }

        public static void ExecuteSync(ExcelFunction function, MethodInfo method, object[] args,
            ref IntPtr result)
        {
            var value = method.Invoke(null, args);
            Converter.ConvertOutging(value, ref result);
        }

        public static void ExecuteAsync(ExcelFunction function, MethodInfo method, object[] args,
            ref IntPtr result)
        {
            Task.Run(() =>
            {
                XLOPER12_num x;
                x.xltype = 1;
                x.num = 0;
                using (var ptr = new StructIntPtr<XLOPER12_num>(ref x))
                {
                    var value = ptr.Ptr;
                    ExecuteSync(function, method, args, ref value);
                    value = ptr.Detach(); // result owned by Excel
                    XlCall.AsyncReturn(function.Index, value);
                }
            });
        }
    }
}
