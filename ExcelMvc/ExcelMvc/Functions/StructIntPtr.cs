using System;
using System.Runtime.InteropServices;

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

        public static T PtrToStruct(IntPtr ptr)
        {
           return Marshal.PtrToStructure<T>(ptr);
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

            Marshal.DestroyStructure<T>(Ptr);
            Marshal.FreeHGlobal(Ptr);
        }
    }
}
