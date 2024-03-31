using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xaml;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        private readonly IntPtr DoubleValue;
        private IntPtr StringValue;
        private readonly IntPtr IntValue;
        private readonly IntPtr ShortValue;

        // thread affinity for return pointers...
        private readonly static ThreadLocal<XlMarshalContext> ThreadInstance
            = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetThreadInstance() => ThreadInstance.Value;

        public XlMarshalContext()
        {
            DoubleValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(double)));
            StringValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(sizeof(short)));
            IntValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));
            ShortValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(short)));

            DoubleToIntPtr(0);
            IntToIntPtr(0);
            ShortToIntPtr(0);
            StringToIntPtr("");
        }

        ~XlMarshalContext()
        {
            Marshal.FreeCoTaskMem(DoubleValue);
            Marshal.FreeCoTaskMem(StringValue);
            Marshal.FreeCoTaskMem(IntValue);
            Marshal.FreeCoTaskMem(ShortValue);
        }
    }
}
