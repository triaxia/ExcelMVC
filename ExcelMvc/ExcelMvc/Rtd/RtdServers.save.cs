using Addin.Interfaces;
using System;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace ExcelMvc.Rtd
{
    internal static class RtdServers__
    {
        private static int nextId = 0;
        private static RtdServer StaticCreate(IRtdServerImpl impl)
        {
            var id = Interlocked.Increment(ref nextId);
            var name = $"ExcelMvc.Rtd.Rtd{id:000}";
            var type = Type.GetType(name);
            var progId = ExcelMvc.Rtd.RtdRegistration.RegisterType(type);
            type = Type.GetTypeFromProgID(progId, true);
            var instance = (RtdServer)Activator.CreateInstance(type);
            //instance.Impl = impl;
            ExcelMvc.Rtd.RtdRegistration.DeleteProgId(progId);
            return instance;
        }

        /// <summary>
        /// Dynamically generated assembly cannot be registered...
        /// </summary>
        /// <param name="impl"></param>
        /// <returns></returns>
        private static RtdServer DynamicCreate(IRtdServerImpl impl)
        {
            var id = Interlocked.Increment(ref nextId);
            var name = $"Rtd{id:000}";
            var cls = CreateClass(name, $"ExcelMvc.{name}");
            CreateConstructor(cls);
            Type type = cls.CreateType();
            var progId = ExcelMvc.Rtd.RtdRegistration.RegisterType(type);
            type = Type.GetTypeFromProgID(progId, true);
            var instance = (RtdServer)Activator.CreateInstance(type);
            //instance.Impl = impl;
            ExcelMvc.Rtd.RtdRegistration.DeleteProgId(progId);
            return instance;
        }

        private static TypeBuilder CreateClass(string className, string progid)
        {
            var asmName = $"{nameof(RtdServers)}.{className}";
            var asmBuilder = AssemblyBuilder.DefineDynamicAssembly(new AssemblyName(asmName), AssemblyBuilderAccess.Run);
            var moduleBuilder = asmBuilder.DefineDynamicModule("MainModule");
            var typeBuilder = moduleBuilder.DefineType(className,
                                TypeAttributes.Public |
                                TypeAttributes.Class |
                                TypeAttributes.AutoClass |
                                TypeAttributes.AnsiClass |
                                TypeAttributes.BeforeFieldInit |
                                TypeAttributes.AutoLayout,
                                typeof(RtdServer));

            AddAtribute(typeBuilder, typeof(ProgIdAttribute)
                , new (Type type, object value)[] { (typeof(string), progid) });
            AddAtribute(typeBuilder, typeof(GuidAttribute)
                , new (Type type, object value)[] { (typeof(string), Guid.NewGuid().ToString()) });
            AddAtribute(typeBuilder, typeof(ComVisibleAttribute)
                , new (Type type, object value)[] { (typeof(bool), true) });
            return typeBuilder;
        }

        private static void CreateConstructor(TypeBuilder typeBuilder)
        {
            typeBuilder.DefineDefaultConstructor(MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName);
        }

        private static void AddAtribute(TypeBuilder typeBuilder, Type attribute, (Type type, object value)[] args)
        {
            var ci = attribute.GetConstructor(args.Select(x => x.type).ToArray());
            var builder = new CustomAttributeBuilder(ci, args.Select(x => x.value).ToArray());
            typeBuilder.SetCustomAttribute(builder);
        }

    }
}
