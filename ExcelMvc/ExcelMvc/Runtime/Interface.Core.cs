#if NET5_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;

namespace ExcelMvc.Runtime
{
    public partial class Interface
    {
        public static int Attach(IntPtr arg, int args)
        {
            //Attach();
            RunInDefaultContext(nameof(Attach));
            return 1;
        }

        public static int Detach(IntPtr arg, int args)
        {
            //Detach();
            RunInDefaultContext(nameof(Detach));
            return 1;
        }

        public static int Click(IntPtr arg, int args)
        {
            Click();
            RunInDefaultContext(nameof(Click));
            return 1;
        }

        public static int Show(IntPtr arg, int args)
        {
            //Show();
            RunInDefaultContext(nameof(Show));
            return 1;
        }

        public static int Hide(IntPtr arg, int args)
        {
            //Hide();
            RunInDefaultContext(nameof(Hide));
            return 1;
        }
        public static int Run(IntPtr arg, int args)
        {
            //Run();
            RunInDefaultContext(nameof(Run));
            return 1;
        }

        private static readonly Dictionary<string, MethodInfo> Methods
            = new Dictionary<string, MethodInfo>()
            {
                {nameof(Attach), null},
                {nameof(Detach), null},
                {nameof(Click), null},
                {nameof(Show), null},
                {nameof(Hide), null},
                {nameof(Run), null}
            };

        static Interface()
        {
            AssemblyLoadContext.Default.Resolving += Default_Resolving;
            var type = typeof(Interface);
            var asm = LoadAssembly(AssemblyLoadContext.Default, type.Assembly.GetName());
            type = asm.GetTypes().Single(x => x.FullName == type.FullName);
            foreach (var name in Methods.Keys)
            {
                var method = type.GetMethod(name, BindingFlags.Public | BindingFlags.Static
                    , Array.Empty<Type>());
                Methods[name] = method;
            }
        }

        public static void RunInDefaultContext(string name, params object[] args)
        {
            Methods[name].Invoke(null, args);
        }

        private static Assembly Default_Resolving(AssemblyLoadContext context, AssemblyName name)
        {
            return LoadAssembly(context, name);
        }

        private static Assembly LoadAssembly(AssemblyLoadContext context, AssemblyName name)
        {
            var asm = context.Assemblies.SingleOrDefault(x => x.FullName == name.FullName);
            if (asm != null) return asm;

            var basePath = System.IO.Path.GetDirectoryName(typeof(Interface).Assembly.Location);
            var folders = context.Assemblies
                .Where(x => !x.IsDynamic && !string.IsNullOrWhiteSpace(x.Location))
                .Select(x => System.IO.Path.GetDirectoryName(x.Location))
                .Concat(new[] { basePath })
                .Distinct();

            var file = folders.Select(x => System.IO.Path.Combine(x!, $"{name.Name}.dll"))
                .Where(System.IO.File.Exists)
                .OrderByDescending(System.IO.File.GetLastWriteTimeUtc)
                .FirstOrDefault();
            return file == null ? null : context.LoadFromAssemblyPath(file);
        }
    }
}
#endif
