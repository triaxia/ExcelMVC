/*
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
        public static void RegisterFunctions()
        {
            var items = DiscoverFunctions().ToArray();
            for (int index = 0; index < items.Length; index++)
                items[index].function.Callback = MakeCallback(items[index].method, items[index].function);

            FunctionHost.Instance.RegisterFunctions(new FunctionDefinitions(items.Select(x => x.function).ToArray()));
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
                .Select(x => (x.type, method: x.type.GetMethod(x.method)))
                .Select(x => (function: x.method.GetCustomAttribute(FunctionHost.Instance.FunctionAttributeType) as IFunctionAttribute, x.method))
                .Select(x => (x.method, (IFunctionAttribute)x.function, GetArguments(x.method)));
        }

        private static IEnumerable<string> GetTypes(Assembly asm)
        {
            return asm.GetExportedTypes().Select(t => (type: t, methods: t.GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(m => m.HasCustomAttribute(FunctionHost.Instance.FunctionAttributeType))))
                .SelectMany(t => t.methods.Select(m => $"{t.type.AssemblyQualifiedName}|{m.Name}"));
        }

        private static ArgumentDefinition[] GetArguments(MethodInfo method)
        {
            return method.GetParameters()
                .Select(x => (argument: x.GetCustomAttribute(FunctionHost.Instance.ArgumentAttributeType) as IArgumentAttribute, parameter: x))
                .Select(x => new ArgumentDefinition(x.parameter, x.argument))
                .ToArray();
        }

        private static bool HasCustomAttribute(this MethodInfo method, Type attributeType)
        {
            var name = attributeType.AssemblyQualifiedName;
            return method.GetCustomAttributesData().Where(x => x.AttributeType.AssemblyQualifiedName == name).Any();
        }

        public static IntPtr MakeCallback(MethodInfo method, FunctionDefinition function)
        {
            var e = DelegateFactory.MakeOuterDelegate(method, function);
            AddIn.NoGarbageCollectableHandles.Add(GCHandle.Alloc(e));
            return Marshal.GetFunctionPointerForDelegate(e);
        }
    }
}
