
/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Developer:         Wolfgang Stamm, Germany

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

namespace ExcelMvc.Runtime
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using ExcelMvc.Diagnostics;
    using Extensions;
#if NET6_0_OR_GREATER
    using System.Runtime.Loader;
#endif
    /// <summary>
    /// Generic object factory
    /// </summary>
    /// <typeparam name="T">Type of object</typeparam>
    public static class ObjectFactory<T>
    {
        private static List<T> Instances { get; } = new List<T>();
        private static bool EqualsNoCase(string lhs, string rhs)
            => StringComparer.InvariantCultureIgnoreCase.Equals(lhs, rhs);

        /// <summary>
        /// Create instances of type T in the current AppDomain.
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void CreateAll()
        {
            var types = GetTypes(out var context);
            FreeReference(context);

            Instances.Clear();
            foreach (var type in types)
            {
                ActionExtensions.Try(() =>
                {
                    var obj = (T)Activator.CreateInstance(Type.GetType(type));
                    Instances.Add(obj);
                }, ex => Messages.Instance.AddErrorLine(ex));
            }
        }

        public static List<string> GetTypes(out WeakReference context)
        {
            List<string> types = new List<string>();
            try
            {
                LoadContext();
                var asms = AppDomain.CurrentDomain.GetAssemblies()
                    .Where(x => !x.IsDynamic);
#if NET6_0_OR_GREATER
                asms = asms.Where(x => !x.IsCollectible);
#endif
                types.AddRange(asms.SelectMany(x => GetTypes(x)));
                var location = typeof(ObjectFactory<T>).Assembly.Location;
                if (!string.IsNullOrWhiteSpace(location))
                {
                    var path = Path.GetDirectoryName(location);
                    var files = Directory.GetFiles(path, "*.dll", SearchOption.TopDirectoryOnly)
                        .Where(x => asms.All(y => !EqualsNoCase(y.Location, x)));
                    var dllTypes = files.SelectMany(x => DiscoverTypes(x));
                    types.AddRange(dllTypes);
                }
            }
            finally
            {
                UnloadContext();
            }
#if NET6_0_OR_GREATER
            context = new WeakReference(AssemblyContext);
#else
            context = null;
#endif
            return types.Distinct().ToList();
        }

        /// <summary>
        /// Deletes instance created
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
        /// Finds the instance matching the full type name specified
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

        private static IEnumerable<string> DiscoverTypes(string assemblyPath)
        {
            var types = Enumerable.Empty<string>();
            ActionExtensions.Try(() =>
            {
                var asm = LoadFrom(assemblyPath);
                types = types.Concat(GetTypes(asm));
            }, ex => Messages.Instance.AddErrorLine(new FileLoadException(ex.Message, assemblyPath, ex)));
            return types;
        }

        private static IEnumerable<string> GetTypes(Assembly asm)
        {
            var itype = typeof(T);
            return asm.GetTypes()
                .Where(x => !x.IsGenericType && !x.IsInterface && !x.IsAbstract && IsDerivedFrom(x, itype))
                .Where(x => x.GetConstructor(Type.EmptyTypes) != null)
                .Select(x => x.AssemblyQualifiedName);
        }

        private static bool IsDerivedFrom(Type type, Type baseType)
        {
            bool IsEqual(Type lhs, Type rhs)
                => (lhs?.AssemblyQualifiedName ?? "") == (rhs?.AssemblyQualifiedName ?? "");
            return IsEqual(type, baseType)
                || IsEqual(type.BaseType, baseType)
                || type.GetInterfaces().Any(x => IsEqual(x, baseType) || IsDerivedFrom(x, baseType));
        }

#if NET6_0_OR_GREATER
        private static AssemblyLoadContext AssemblyContext {get; set;}
#endif
        private static void LoadContext()
        {
#if NET6_0_OR_GREATER
            UnloadContext();
            AssemblyContext = new AssemblyLoadContext($"ObjectFactory<{typeof(T)}>", true);
            AssemblyContext.Resolving +=(sender, args) =>
            {
                var file = sender.Assemblies.Where(x => !x.IsDynamic)
                    .Select(x => Path.GetDirectoryName(x.Location))
                    .Distinct()
                    .Select(x => Path.ChangeExtension(Path.Combine(x!, args.Name!), ".dll"))
                    .Where(File.Exists)
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();
                return file == null ? null : sender.LoadFromAssemblyPath(file);
            };
#else
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve +=
                (_, args) => Assembly.ReflectionOnlyLoad(args.Name);
#endif
        }
        private static void UnloadContext()
        {
#if NET6_0_OR_GREATER
            AssemblyContext?.Unload();
            AssemblyContext = null;
#endif
        }
        private static Assembly LoadFrom(string assemblyPath)
        {
#if NET5_0_OR_GREATER
            var loaded = AssemblyContext.Assemblies
                .SingleOrDefault(a => !a.IsDynamic && EqualsNoCase(a.Location, assemblyPath));
            return loaded ?? AssemblyContext.LoadFromAssemblyPath(assemblyPath);
#else
            return Assembly.ReflectionOnlyLoadFrom(assemblyPath);
#endif
        }

        private static void FreeReference(WeakReference reference)
        {
            if (reference == null) return;

            var timeout = TimeSpan.FromSeconds(10);
            var start = DateTime.UtcNow;
            while (reference.IsAlive && (DateTime.UtcNow - start) < timeout)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (reference.IsAlive)
                throw new TimeoutException($"ObjectFactory<{typeof(T)}>.CreateAll timed out {timeout}");
        }
    }
}
