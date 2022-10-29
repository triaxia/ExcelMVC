using Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMvc.Functions
{
    public class StructIntPtr<T> : IDisposable
    {
        public IntPtr Ptr { get; private set; } = IntPtr.Zero;

        public StructIntPtr(ref T structure)
        {
            Ptr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(T)));
            Marshal.StructureToPtr(structure, Ptr, false);
        }

        public static implicit operator IntPtr(StructIntPtr<T> me)
        {
            return me.Ptr;
        }

        public IntPtr Detach()
        {
            var me = Ptr;
            Ptr = IntPtr.Zero;
            return me;
        }

        public void Dispose()
        {
            if (Ptr == IntPtr.Zero)
                return;

            Marshal.DestroyStructure<Function>(Ptr);
            Marshal.FreeHGlobal(Ptr);
        }
    }
}
