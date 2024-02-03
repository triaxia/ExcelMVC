#if NET5_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;

namespace ExcelMvc.Runtime
{
    public static partial class Interface
    {
        public static int Attach(IntPtr arg, int args)
        {
            Attach(arg);
            return 1;
        }

        public static int Detach(IntPtr arg, int args)
        {
            Detach();
            return 1;
        }

        public static int Click(IntPtr arg, int args)
        {
            Click();
            return 1;
        }

        public static int Show(IntPtr arg, int args)
        {
            Show();
            return 1;
        }

        public static int Hide(IntPtr arg, int args)
        {
            Hide();
            return 1;
        }
        public static int Run(IntPtr arg, int args)
        {
            Run();
            return 1;
        }
    }
}
#endif
