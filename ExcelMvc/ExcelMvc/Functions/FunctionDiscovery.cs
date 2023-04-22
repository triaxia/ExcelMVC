using ExcelMvc.Runtime;
using Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class FunctionDiscovery
    {
        public static IEnumerable<(MethodInfo method, FunctionAttribute function, Argument[] args)> Discover()
        {
            return ObjectFactory<object>.GetTypes(x => GetTypes(x), ObjectFactory<object>.SelectAllAssembly)
                .Select(x => x.Split('|')).Select(x => (type: Type.GetType(x[0]), method: x[1]))
                .Select(x => (x.type, method: x.type.GetMethod(x.method)))
                .Select(x => (function: x.method.GetCustomAttribute<FunctionAttribute>(), x.method))
                .Select(x => (x.method, (FunctionAttribute)x.function, GetArguments(x.method)));
        }

        private static IEnumerable<string> GetTypes(Assembly asm)
        {
            return asm.GetTypes().Select(t => (type: t, methods: t.GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(m => m.HasCustomAttribute<FunctionAttribute>())))
                .SelectMany(t => t.methods.Select(m => $"{t.type.AssemblyQualifiedName}|{m.Name}"));
        }

        private static Argument[] GetArguments(MethodInfo method)
        {
            return method.GetParameters()
                .Select(x => (argument: x.GetCustomAttribute<ArgumentAttribute>(), parameter: x))
                .Select(x => x.argument == null ? new Argument { Name = x.parameter.Name, Description = "" } : new Argument(x.argument))
                .ToArray();
        }

        private static bool HasCustomAttribute<T>(this MethodInfo method) where T : Attribute
        {
            var name = typeof(T).AssemblyQualifiedName;
            return method.GetCustomAttributesData().Where(x => x.AttributeType.AssemblyQualifiedName == name).Any();
            //return method.GetCustomAttributes().Where(x => x.GetType().AssemblyQualifiedName == name).Any();
        }
    }
}
