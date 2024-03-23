using ExcelMvc.Interfaces;
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

        public RtdRegistry(Type implType, Func<IRtdServerImpl> implFactory)
        {
            var key = implType.FullName;
            var pair = Servers.GetOrAdd(key, _ =>
            {
                Registered = true;
                var (progId, guid) = RegistryFunctions.Register();
                var impl = implFactory?.Invoke() ?? (IRtdServerImpl)Activator.CreateInstance(implType);
                var server = new RtdServer(impl);
                Factories[guid] = new RtdComClassFactory(server);
                return (new RtdServer(impl), progId);
            });
            ProgId = pair.progId;
        }

        public void Dispose()
        {
            if (Registered)
                RegistryFunctions.Unregister(ProgId);
            GC.SuppressFinalize(this);
        }

        public static RtdComClassFactory FindFactory(Guid guid)
        {
            return Factories.ToArray().SingleOrDefault(x => x.Key == guid).Value;
        }

        public static void OnTerminated(RtdServer server)
        {
            var key = Servers.ToArray().SingleOrDefault(x => x.Value.server == server).Key;
            Servers.TryRemove(key, out var _);
            var guid = Factories.ToArray().SingleOrDefault(x => x.Value.RtdServer == server).Key;
            Factories.TryRemove(guid, out var _);
        }
    }
}
