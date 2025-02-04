﻿/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor,
Boston, MA 02110-1301 USA.
*/

using ExcelMvc.Runtime;
using Function.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class FunctionDiscovery
    {
        public static FunctionDefinitions RegisterFunctions()
        {
            var items = DiscoverFunctions().ToArray();
            for (int index = 0; index < items.Length; index++)
                items[index].function.Callback = MakeCallback(items[index].method, items[index].function);

            var functions = new FunctionDefinitions(items.Select(x => x.function).ToArray());
            FunctionHost.Instance.RegisterFunctions(functions);
            return functions;
        }

        public static IEnumerable<(MethodInfo method, FunctionDefinition function)> DiscoverFunctions()
        {
            return Discover()
                .Select(x => (x.method, x.function, x.args))
                .Select(x => (x.method, function: new FunctionDefinition(x.function, x.args, IntPtr.Zero, x.method)))
                .ToArray();
        }

        private static IEnumerable<(MethodInfo method, IFunctionAttribute function, ArgumentDefinition[] args)> Discover()
        {
            return ObjectFactory<object>.GetTypes(x => GetTypes(x), ObjectFactory<object>.SelectAllAssembly)
                .Select(x => x.Split('|')).Select(x => (type: Type.GetType(x[0]), method: x[1]))
                .Select(x => (MatchFirst(x.type, x.method)))
                .Select(x => (x.method, x.function, GetArguments(x.method)));
        }

        private static (IFunctionAttribute function, MethodInfo method)
            MatchFirst(Type type, string name)
        {
            // discard duplicate functions...
            return type.GetMethods(BindingFlags.Public | BindingFlags.Static).Where(x => x.Name == name)
                .Select(x => (function: GetFunctionAttribute(x), method: x))
                .Where(x => x.function != null).FirstOrDefault();
        }

        private static IEnumerable<string> GetTypes(Assembly asm)
        {
            return asm.GetExportedTypes().Select(t => (type: t, methods: t.GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(m => HasFunctionAttribute(m))))
                .SelectMany(t => t.methods.Select(m => $"{t.type.AssemblyQualifiedName}|{m.Name}"));
        }

        private static ArgumentDefinition[] GetArguments(MethodInfo method)
        {
            return method.GetParameters()
                .Select(x => (argument: GetArgumentAttribute(x), parameter: x))
                .Select(x => new ArgumentDefinition(x.parameter, x.argument))
                .ToArray();
        }

        private static bool HasFunctionAttribute(MethodInfo method)
        {
            // DON'T use GetFunctionAttribute(method) as it will return null because the method is
            // NOT a fully reflectable one yet...
            var name = typeof(IFunctionAttribute).FullName;
            return method.GetCustomAttributesData()
                .Where(x => x.AttributeType.GetInterfaces().Any(y => y.FullName == name))
                .Any();
        }

        private static IFunctionAttribute GetFunctionAttribute(MethodInfo method)
            => method.GetCustomAttributes().SingleOrDefault(x => x is IFunctionAttribute) as IFunctionAttribute;

        private static IArgumentAttribute GetArgumentAttribute(ParameterInfo parameter)
            => parameter.GetCustomAttributes().SingleOrDefault(x => x is IArgumentAttribute) as IArgumentAttribute;

        private static IntPtr MakeCallback(MethodInfo method, FunctionDefinition function)
        {
            var e = DelegateFactory.MakeOuterDelegate(method, function);
            AddIn.NoGarbageCollectableHandles.Add(GCHandle.Alloc(e));
            return Marshal.GetFunctionPointerForDelegate(e);
        }
    }
}
