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
            var fargs = Marshal.PtrToStructure<FunctionArgs>(args);
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

        public static FunctionCallback MakeCallback(MethodInfo method, FunctionAttribute function)
        {
            return new FunctionCallback(FunctionExecution.Execute);
        }

        /* 9 ms version 300 ms for 1million calls
         static void Main(string[] args)
        {
            var ps = new int[] { 1, 2 }.Select((idx, x) => Expression.Parameter(typeof(int), $"a{idx}")).ToArray();
            var m = typeof(Program).GetMethod("X")!;
            var d = Expression.Call(typeof(Program).GetMethod("X")!, ps);
            var f = ((Func<int, int, int>)(Expression.Lambda(d, ps).Compile()));
            //var f = Expression.Lambda(d, ps).Compile();

            var start = System.Diagnostics.Stopwatch.StartNew();
            for (var idx = 0; idx < 1000000; idx++)
            {
                var v = f(3, 4);
            }
            Console.WriteLine(start.Elapsed.ToString());

            start = System.Diagnostics.Stopwatch.StartNew();
            for (var idx = 0; idx < 1000000; idx++)
            {
                var v = m.Invoke(null, new object[] { 1, 3 });
            }
            Console.WriteLine(start.Elapsed.ToString());
        }

     static void Main(string[] args)
        {
            var p1 = new int[] { 1, 2 }.Select((idx, x) => Expression.Parameter(typeof(int), $"a{idx}")).ToArray();
            var m1 = typeof(Program).GetMethod("X")!;
            var c1 = Expression.Call(m1, p1);
            var f1 = ((Func<int, int, int>)(Expression.Lambda(c1, p1).Compile()));

            var p2 = new[] { Expression.Parameter(typeof(object), "c"), Expression.Parameter(typeof(Type), "p") };
            var m2 = typeof(Program).GetMethod("C")!;
            var c2 = Expression.Call(m2, p2);
            var f2 = ((Func<object, Type, int>)(Expression.Lambda(c2, p2).Compile()));

        var start = System.Diagnostics.Stopwatch.StartNew();
            for (var idx = 0; idx< 1000000; idx++)
            {
                //var v = f1(f2(3, typeof(int)), f2(4, typeof(int)));
                var v = f1(3, 4);
    }
    Console.WriteLine(start.Elapsed.ToString());

            start = System.Diagnostics.Stopwatch.StartNew();
            for (var idx = 0; idx< 1000000; idx++)
            {
                //var v = m1.Invoke(null, new object[] { (1, typeof(int)), C(3, typeof(int))});
                var v = m1.Invoke(null, new object[] { 1, 3 });
}
Console.WriteLine(start.Elapsed.ToString());
        }

        */
    }
}
