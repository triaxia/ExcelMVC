using ExcelMvc.Rtd;
using Addin.Interfaces;
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
            var fargs = Marshal.PtrToStructure<FunctionArgs>(args);
            var (method, function, callback) = Functions[fargs.Index];
            var argc = method.GetParameters().Length;
            var arguments = fargs.GetArgs(argc);

            var values = method.GetParameters()
                .Select((p, idx) => Converter.FromIncoming(XLOPER12.FromIntPtr(arguments[idx]), p))
                .ToArray();

            if (function.IsAsync)
                ExecuteAsync(function, method, values, fargs.GetArgs()[argc], ref fargs.Result);
            else
                ExecuteSync(function, method, values, ref fargs.Result);
        }

        public static void ExecuteSync(Function function, MethodInfo method, object[] args,
            ref IntPtr result)
        {
            var value = method.Invoke(null, args);
            Converter.ToOutgoing(value, ref result, method);
        }

        public static void ExecuteAsync(Function function, MethodInfo method, object[] args
            , IntPtr handle, ref IntPtr result)
        {
            Converter.ToOutgoing("...", ref result, method);
            Task.Factory.StartNew(state =>
            {
                var largs = (object[])state;
                var outcome = XLOPER12.FromObject(method.Invoke(null, (object[])largs[2]));
                using (var presult = new StructIntPtr<XLOPER12>(ref outcome))
                    XlCall.AsyncReturn((IntPtr)largs[3], presult.Detach());

            }, new object[] { function, method, args, handle });
        }

        public static object ExecuteRtd()
        {
            var (type, progId) = RtdServers.Acquire(new RtdServerImplTest());
            FunctionArgs args = new FunctionArgs();
            var x = XLOPER12.FromObject(progId);
            var y = XLOPER12.FromObject("");
            var z = XLOPER12.FromObject("");
            using (var xx = new StructIntPtr<XLOPER12>(ref x))
            using (var yy = new StructIntPtr<XLOPER12>(ref y))
            using (var zz = new StructIntPtr<XLOPER12>(ref z))
            {
                args.Arg00 = xx.Ptr;
                args.Arg01 = yy.Ptr;
                args.Arg02 = zz.Ptr;
                using (var p = new StructIntPtr<FunctionArgs>(ref args))
                {
                    var result = XLOPER12.FromIntPtr(XlCall.RtdCall(p.Ptr));
                    return result == null ? null : XLOPER12.ToObject(result.Value);
                }
            }
        }

        public static FunctionCallback MakeCallback(MethodInfo method, FunctionAttribute function)
        {
            return new FunctionCallback(FunctionExecution.Execute);
        }
    }
}
