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

using Function.Interfaces;
using System;
using System.Collections.Concurrent;
using System.Linq;
using static ExcelMvc.Rtd.RtdServerFactory;

namespace ExcelMvc.Rtd
{
    public class RtdRegistry : IDisposable
    {
        public string ProgId { get; }

        private static readonly ConcurrentDictionary<string, (RtdServer server, string progId)> Servers
            = new ConcurrentDictionary<string, (RtdServer server, string progId)>();
        private static readonly ConcurrentDictionary<Guid, RtdComClassFactory> Factories
            = new ConcurrentDictionary<Guid, RtdComClassFactory>();

        private bool Registered { get; set; }

        public RtdRegistry(Type type, Func<IRtdServerImpl> factory)
        {
            lock (Servers)
            {
                RegistryFunctions.PurgeProgIds();
                var key = type.FullName;
                var pair = Servers.GetOrAdd(key, _ =>
                {
                    Registered = true;
                    var (progId, guid) = RegistryFunctions.Register();
                    var impl = factory?.Invoke() ?? (IRtdServerImpl)Activator.CreateInstance(type);
                    var server = new RtdServer(impl);
                    Factories[guid] = new RtdComClassFactory(server);
                    return (server, progId);
                });
                ProgId = pair.progId;
            }
        }

        public void Dispose()
        {
            /*
            if (Registered)
                RegistryFunctions.Unregister(ProgId);
            */
        }

        public static RtdComClassFactory FindFactory(Guid guid)
        {
            lock (Servers)
            {
                return Factories.ToArray().SingleOrDefault(x => x.Key == guid).Value;
            }
        }

        public static void OnTerminated(RtdServer server)
        {
            lock (Servers)
            {
                var key = Servers.ToArray().SingleOrDefault(x => x.Value.server == server).Key;
                Servers.TryRemove(key, out var _);
                var guid = Factories.ToArray().SingleOrDefault(x => x.Value.RtdServer == server).Key;
                Factories.TryRemove(guid, out var _);
            }
        }
    }
}
