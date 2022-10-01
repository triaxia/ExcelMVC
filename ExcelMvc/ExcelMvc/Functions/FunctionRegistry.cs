using ExcelMvc.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class FunctionRegistry
    {
        public static List<(ExcelFunction function, MethodInfo method)> Functions { get; private set; } 
        public static void Register()
        {
            Functions = Discover().ToList();
            foreach (var (function, _) in Functions)
                XlCall.Register(function);
        }

        public static IEnumerable<(ExcelFunction function, MethodInfo method)> Discover()
        {
            return ObjectFactory<object>.GetTypes(x => GetTypes(x), ObjectFactory<object>.SelectAllAssembly)
                .Select(x => x.Split('|')).Select(x => (type: Type.GetType(x[0]), method: x[1]))
                .Select(x => (x.type, method: x.type.GetMethod(x.method)))
                .Select(x => (function: x.method.GetCustomAttribute<ExcelFunctionAttribute>(), x.method))
                .Select((x, idx) => (new ExcelFunction((uint)idx, (ExcelFunctionAttribute)x.function, GetArguments(x.method)), x.method));
        }

        private static IEnumerable<string> GetTypes(Assembly asm)
        {
            return asm.GetTypes().Select(t => (type: t, methods: t.GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(m => m.HasCustomAttribute<ExcelFunctionAttribute>())))
                .SelectMany(t => t.methods.Select(m => $"{t.type.AssemblyQualifiedName}|{m.Name}"));
        }

        private static ExcelArgument[] GetArguments(MethodInfo method)
        {
            return method.GetParameters()
                .Select(x => (argument: x.GetCustomAttribute<ExcelArgumentAttribute>(), parameter: x))
                .Select(x => x.argument == null ? new ExcelArgument { Name = x.parameter.Name, Description = "" } : new ExcelArgument(x.argument))
                .ToArray();
        }

        private static bool HasCustomAttribute<T>(this MethodInfo method) where T : Attribute
        {
            var name = typeof(T).AssemblyQualifiedName;
            return method.GetCustomAttributes().Where(x => x.GetType().AssemblyQualifiedName == name).Any();
        }
    }
}
