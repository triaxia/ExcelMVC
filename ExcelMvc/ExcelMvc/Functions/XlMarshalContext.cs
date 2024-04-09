using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        private readonly IntPtr DoubleValue;
        private IntPtr StringValue = IntPtr.Zero;
        private readonly IntPtr IntValue;
        private readonly IntPtr ShortValue;
        private IntPtr DoubleArrayValue = IntPtr.Zero;

        // thread affinity for return pointers...
        private readonly static ThreadLocal<XlMarshalContext> ThreadInstance
            = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetThreadInstance() => ThreadInstance.Value;

        public XlMarshalContext()
        {
            DoubleValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(double)));
            IntValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));
            ShortValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(short)));

            DoubleToIntPtr(0);
            Int32ToIntPtr(0);
            Int16ToIntPtr(0);
        }

        ~XlMarshalContext()
        {
            Marshal.FreeCoTaskMem(DoubleValue);
            Marshal.FreeCoTaskMem(StringValue);
            Marshal.FreeCoTaskMem(IntValue);
            Marshal.FreeCoTaskMem(ShortValue);
            Marshal.FreeCoTaskMem(DoubleArrayValue);
        }
    }
}
