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
        public static Dictionary<uint, (MethodInfo method, ExcelFunction function, FunctionCallback callback)> Functions
        { get; private set; }

        public static void RegisterFunctions()
        {
            Functions = FunctionDiscovery.Discover()
                .Select((x, idx) => (index: (uint)idx, x.method, x.function, x.args, callback: MakeCallback(x.method, x.function)))
                .ToDictionary(x => x.index, x => (x.method, new ExcelFunction(x.index, x.function, x.args, x.callback), x.callback));
            foreach (var pair in Functions)
                XlCall.RegisterFunction(pair.Value.function);
        }

        public static void Execute(IntPtr args)
        {
            var fargs = Marshal.PtrToStructure<FunctionArgs>(args);
            var (method, function, callback) = Functions[fargs.Index];
            var arguments = fargs.GetArgs();

            var argc = method.GetParameters().Length;
            var values = method.GetParameters().Take(argc)
                .Select((p, idx) => Converter.ConvertIncoming(arguments[idx], p))
                .ToArray();

            if (function.IsAnyc)
                ExecuteAsync(function, method, values, fargs.GetArgs()[argc]);
            else
                ExecuteSync(function, method, values, ref fargs.Result);
        }

        public static void ExecuteSync(ExcelFunction function, MethodInfo method, object[] args,
            ref IntPtr result)
        {
            var value = method.Invoke(null, args);
            Converter.ConvertOutging(value, ref result);
        }

        public static void ExecuteAsync(ExcelFunction function, MethodInfo method, object[] args, IntPtr handle)
        {
            Task.Factory.StartNew(state =>
            {
                var largs = (object[])state;
                XLOPER12_num x;
                x.xltype = 1;
                x.num = 0;
                using (var result = new StructIntPtr<XLOPER12_num>(ref x))
                {
                    var value = result.Ptr;
                    ExecuteSync((ExcelFunction)largs[0], (MethodInfo)largs[1], (object[])largs[2], ref value);
                    // result will be owned and freed by Excel, so detach it.
                    XlCall.AsyncReturn((IntPtr)largs[3], result.Detach());
                }
            }, new object[] { function, method, args, handle });
        }

        public static FunctionCallback MakeCallback(MethodInfo method, ExcelFunctionAttribute function)
        {
            return new FunctionCallback(FunctionExecution.Execute);
        }
    }
}
