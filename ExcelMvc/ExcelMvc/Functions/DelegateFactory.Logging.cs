using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public class ExecutingEventArgs : EventArgs
    {
        public string Name { get; }
        public MethodInfo Method { get; }
        public object[] Args { get; }
        public ExecutingEventArgs(string name, MethodInfo method, object[] args)
        {
            Name = name;
            Method = method;
            Args = args;
        }

        public override string ToString()
        {
            var args = string.Join(",", Method.GetParameters()
                .Select((p, i) => (name: p.Name, value: $"{Args[i]}"))
                .Select(x => $"{x.name}={x.value}"));
            return $"{Name}[{args}]";
        }
    }

    public static partial class DelegateFactory
    {
        public static event EventHandler<ExecutingEventArgs> Executing;

        public static void Log0(string name, MethodInfo method) => Log(name, method, Array.Empty<object>());
        public static void Log1(string name, MethodInfo method, object a1) => Log(name, method, new object[] { a1 });
        public static void Log2(string name, MethodInfo method, object a1, object a2) => Log(name, method, new object[] { a1, a2 });
        public static void Log3(string name, MethodInfo method, object a1, object a2, object a3) => Log(name, method, new object[] { a1, a2, a3 });

        private static void Log(string name, MethodInfo method, object[] args)
        {
            Executing.Invoke(null, new ExecutingEventArgs(name, method, args));
        }

        private static readonly Dictionary<int, MethodInfo> LoggingMethods
            = new Dictionary<int, MethodInfo>()
            {
                { 0, typeof(DelegateFactory).GetMethod(nameof(Log0)) },
                { 1, typeof(DelegateFactory).GetMethod(nameof(Log1)) },
                { 2, typeof(DelegateFactory).GetMethod(nameof(Log2)) },
                { 3, typeof(DelegateFactory).GetMethod(nameof(Log3)) }
            };

        public static MethodInfo LoggingMethod(int parametersCount) =>
            LoggingMethods.TryGetValue(parametersCount, out var value) ? value : LoggingMethods[0];
    }
}
