using ExcelMvc.Rtd;
using Mvc;
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
        public static Dictionary<int, (MethodInfo method, Function function, FunctionCallback callback)> Functions
        { get; private set; }

        public static void RegisterFunctions()
        {
            Functions = FunctionDiscovery.Discover()
                .Select((x, idx) => (index: idx, x.method, x.function, x.args, callback: MakeCallback(x.method, x.function)))
                .ToDictionary(x => x.index, x => (x.method, new Function(x.index, x.function, x.args, x.callback), x.callback));
            foreach (var pair in Functions)
                XlCall.RegisterFunction(pair.Value.function);
        }

        public static void Execute(IntPtr args)
        {
            var fargs = StructIntPtr<FunctionArgs>.PtrToStruct(args);
            var (method, function, callback) = Functions[fargs.Index];
            var argc = method.GetParameters().Length;
            var arguments = fargs.GetArgs(argc);

            var values = method.GetParameters()
                .Select((p, idx) => Converter.ConvertIncoming(arguments[idx], p))
                .ToArray();

            if (function.IsAsync)
                ExecuteAsync(function, method, values, fargs.GetArgs()[argc]);
            else
                ExecuteSync(function, method, values, ref fargs.Result);
        }

        public static void ExecuteSync(Function function, MethodInfo method, object[] args,
            ref IntPtr result)
        {
            var value = method.Invoke(null, args);
            Converter.ConvertOutging(value, method, ref result);
        }

        public static void ExecuteAsync(Function function, MethodInfo method, object[] args, IntPtr handle)
        {
            Task.Factory.StartNew(state =>
            {
                var largs = (object[])state;
                var r = new XLOPER12((double)0);
                using (var result = new StructIntPtr<XLOPER12>(ref r))
                {
                    var value = result.Ptr;
                    ExecuteSync((Function)largs[0], (MethodInfo)largs[1], (object[])largs[2], ref value);
                    // result will be owned and freed by Excel, so detach it.
                    XlCall.AsyncReturn((IntPtr)largs[3], result.Detach());
                }
            }, new object[] { function, method, args, handle });
        }

        public static object ExecuteRtd()
        {
            RtdRegistration.RegisterType(typeof(Rtd002));
            FunctionArgs args = new FunctionArgs();
            var x = new XLOPER12("ExcelMvc.Rtd002");
            var y = new XLOPER12("");
            var z = new XLOPER12("");
            using (var xx = new StructIntPtr<XLOPER12>(ref x))
            using (var yy = new StructIntPtr<XLOPER12>(ref y))
            using (var zz = new StructIntPtr<XLOPER12>(ref z))
            {
                args.Arg00 = xx.Ptr;
                args.Arg01 = yy.Ptr;
                args.Arg02 = zz.Ptr;
                using (var p = new StructIntPtr<FunctionArgs>(ref args))
                {
                    var result = XlCall.RtdCall(p.Ptr);
                    return Marshal.PtrToStructure<XLOPER12>(result).num;
                }
            }
        }

        public static FunctionCallback MakeCallback(MethodInfo method, FunctionAttribute function)
        {
            return new FunctionCallback(FunctionExecution.Execute);
        }
    }
}
