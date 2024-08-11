
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

#if NET6_0_OR_GREATER
using System.Runtime.Loader;
#endif

namespace ExcelMvc.Runtime
{
    /// <summary>
    /// Resloves assemblies for the current AppDomain or the current AssemblyLoadContext.
    /// </summary>
    public sealed class AssemblyResolver : IDisposable
    {
        private readonly HashSet<string> BasePaths = new HashSet<string>();
#if NET6_0_OR_GREATER
        private AssemblyLoadContext Context { get;}
#endif
        /// <summary>
        /// Initializes a new instance of <see cref="AssemblyResolver"/>
        /// </summary>
        /// <param name="paths"></param>
        public AssemblyResolver(IEnumerable<string> paths)
        {
            foreach (var path in paths)
                BasePaths.Add(path);
#if NET6_0_OR_GREATER
            Context = AssemblyLoadContext.GetLoadContext(Assembly.GetExecutingAssembly());    
            Context.Resolving += AssemblyResolve;
#else
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolve;
#endif
        }

        /// <summary>
        /// <inheritdoc cref="IDisposable.Dispose"/>
        /// </summary>
        public void Dispose()
        {
#if NET6_0_OR_GREATER
            Context.Resolving -= AssemblyResolve;
#else
            AppDomain.CurrentDomain.AssemblyResolve -= AssemblyResolve;
#endif
        }

        /// <summary>
        /// Loads an assembly into the current AppDomain or the current AssemblyLoadContext.
        /// </summary>
        /// <param name="assemblyPath"></param>
        /// <returns></returns>
        public Assembly LoadAssembly(string assemblyPath)
        {
            try
            {
                var name = AssemblyName.GetAssemblyName(assemblyPath).FullName;
#if NET6_0_OR_GREATER
                var match = Context.Assemblies
                    .Where(x => !x.IsDynamic && x.GetName().FullName == name)
                    .SingleOrDefault();
                return match ?? Context.LoadFromAssemblyPath(assemblyPath);
#else
                var match = AppDomain.CurrentDomain.GetAssemblies()
                    .Where(x => !x.IsDynamic && x.GetName().FullName == name)
                    .SingleOrDefault();
                return match ?? Assembly.LoadFrom(assemblyPath);
#endif
            }
            catch (BadImageFormatException)
            {
                return null;
            }
        }

#if NET6_0_OR_GREATER
        private Assembly AssemblyResolve(AssemblyLoadContext _, AssemblyName arg2)
        {
            var name = $"{arg2.Name}.dll";
            var match = BasePaths.Select(x => Path.Combine(x, name))
                .Select(x => File.Exists(x) ? x : null)
                .Where(x => x != null)
                .SingleOrDefault();
            return match == null ? null : LoadAssembly(match);
        }
#else
        private Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var name = $"{new AssemblyName(args.Name).Name}.dll";
            var match = BasePaths.Select(x => Path.Combine(x, name))
                .Select(x => File.Exists(x) ? x : null)
                .Where(x => x != null)
                .SingleOrDefault();
            return match == null ? null : LoadAssembly(match);
        }
#endif
    }
}