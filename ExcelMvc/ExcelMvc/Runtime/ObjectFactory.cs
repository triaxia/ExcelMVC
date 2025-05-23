
/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany (2013)

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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelMvc.Extensions;
using Function.Interfaces;

namespace ExcelMvc.Runtime
{
    /// <summary>
    /// Generic object factory.
    /// </summary>
    /// <typeparam name="T">Type of object</typeparam>
    public static class ObjectFactory<T>
    {
        private static bool EqualsIgnoreCase(string lhs, string rhs)
            => StringComparer.InvariantCultureIgnoreCase.Equals(lhs, rhs);
        private static bool StartsWithIgnoreCase(string lhs, string rhs)
            => lhs.StartsWith(rhs, StringComparison.InvariantCultureIgnoreCase);

        /// <summary>
        /// Instances created.
        /// </summary>
        public static List<T> Instances { get; } = new List<T>();

        /// <summary>
        /// Gets and sets the function that selects all assemblies.
        /// </summary>
        public static Func<string, bool, bool> SelectAllAssembly { get; set; }
            = (name, loaded) => !StartsWithIgnoreCase(name, "Microsoft")
                             && !StartsWithIgnoreCase(name, "System");

        /// <summary>
        /// Create instances of type T in the current AppDomain.
        /// </summary>
        /// <param name="getTypes"></param>
        /// <param name="selectAssembly">The function that takes arguments (assembly 
        /// name or file name, loaded or not) and returns true or false to indicate
        /// if an assembly should be included in the discover process</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void CreateAll(Func<Assembly, IEnumerable<string>> getTypes,
            Func<string, bool, bool> selectAssembly)
        {
            Instances.Clear();
            var types = GetTypes(getTypes, selectAssembly);
            foreach (var type in types)
            {
                ActionExtensions.Try(() =>
                {
                    var obj = (T)Activator.CreateInstance(Type.GetType(type));
                    Instances.Add(obj);
                }, ex => FunctionHost.Instance.RaiseFailed(FunctionHost.Instance, new ErrorEventArgs(ex)));
            }
        }

        /// <summary>
        /// Deletes instance created.
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void DeleteAll(Action<T> disposer)
        {
            if (Instances != null)
            {
                if (disposer != null)
                    Instances.ForEach(disposer);
                Instances.Clear();
            }
        }

        /// <summary>
        /// Finds the instance matching the full type name specified.
        /// </summary>
        /// <param name="fullTypeName"></param>
        /// <returns></returns>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static T Find(string fullTypeName)
        {
            var idx = Instances.FindIndex(x => x.GetType().FullName == fullTypeName);
            if (idx < 0)
                idx = Instances.FindIndex(x => x.GetType().AssemblyQualifiedName == fullTypeName);
            return idx < 0 ? default : Instances[idx];
        }

        /// <summary>
        /// Gets the creatable types with default constructors from the 
        /// specified assembly.
        /// </summary>
        /// <param name="asm"></param>
        /// <returns></returns>
        public static IEnumerable<string> GetCreatableTypes(Assembly asm)
        {
            var itype = typeof(T);
            return asm.GetExportedTypes()
                .Where(x => !x.IsGenericType && !x.IsInterface && !x.IsAbstract && IsDerivedFrom(x, itype))
                .Where(x => x.GetConstructor(Type.EmptyTypes) != null)
                .Select(x => x.AssemblyQualifiedName);
        }

        /// <summary>
        /// Discovers types of type T in the current AppDomain.
        /// </summary>
        /// <param name="selectAssembly">A function (assembly name or file name,
        /// loaded or npt) that returns true or false to indicate if an assembly 
        /// should be included in the discover process</param>
        /// <returns></returns>
        public static List<string> GetTypes(Func<Assembly, IEnumerable<string>> getTypes
            , Func<string, bool, bool> selectAssembly)
        {
            var types = new List<string>();
#if NET6_0_OR_GREATER
            var asms = System.Runtime.Loader.AssemblyLoadContext.GetLoadContext(Assembly.GetExecutingAssembly()).Assemblies
                .Where(x=>!x.IsDynamic);
#else
            var asms = AppDomain.CurrentDomain.GetAssemblies().Where(x => !x.IsDynamic);
#endif
            asms = asms.Where(x => selectAssembly(x.GetName().Name, true));
            types.AddRange(asms.SelectMany(x => getTypes(x)));
            var location = typeof(ObjectFactory<T>).Assembly.Location;
            if (!string.IsNullOrWhiteSpace(location))
            {
                var path = Path.GetDirectoryName(location);
                var files = GetAssemblyFiles(path).Where(x => asms.All(y => !EqualsIgnoreCase(y.Location, x)));
                if (selectAssembly != null)
                    files = files.Where(x => selectAssembly(Path.GetFileNameWithoutExtension(x), true));

                if (files.Any())
                {
                    var dllTypes = files.SelectMany(x => DiscoverTypes(x, getTypes));
                    types.AddRange(dllTypes);
                }
            }
            return types.Distinct().ToList();
        }

        private static IEnumerable<string> DiscoverTypes(string assemblyPath,
            Func<Assembly, IEnumerable<string>> getTypes)
        {
            var types = Enumerable.Empty<string>();
            ActionExtensions.Try(() =>
            {
                using (var resolver = new AssemblyResolver(new[] { Path.GetDirectoryName(assemblyPath) }))
                {
                    var asm = resolver.LoadAssembly(assemblyPath);
                    if (asm != null) types = types.Concat(getTypes(asm));
                }
            }, ex => RaiseFailed(new FileLoadException(ex.Message, assemblyPath, ex)));
            return types;
        }

        private static bool IsDerivedFrom(Type type, Type baseType)
        {
            bool IsEqual(Type lhs, Type rhs)
                => (lhs?.AssemblyQualifiedName ?? "") == (rhs?.AssemblyQualifiedName ?? "");
            return IsEqual(type, baseType)
                || IsEqual(type.BaseType, baseType)
                || type.GetInterfaces().Any(x => IsEqual(x, baseType) || IsDerivedFrom(x, baseType));
        }

        private static string[] GetAssemblyFiles(string path)
        {
            var patterns = Directory.GetFiles(path, "*.reflection.txt")
                .SelectMany(x => File.ReadAllLines(x).Select(y => y.Trim()).Where(y => y != "" && !y.StartsWith("#")))
                .Distinct();
            if (!patterns.Any()) patterns = new[] { "*.dll" };
            return patterns.SelectMany(x => Directory.GetFiles(path, x, SearchOption.TopDirectoryOnly))
                .Distinct().ToArray();
        }

        private static void RaiseFailed(Exception ex)
            => FunctionHost.Instance.RaiseFailed(FunctionHost.Instance, new ErrorEventArgs(ex));
    }
}
