using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        private readonly IntPtr DoubleValue;
        private readonly IntPtr DecimalValue;
        private readonly IntPtr StringValue;
        private readonly IntPtr LongValue;

        // thread affinity for return pointers...
        private readonly static ThreadLocal<XlMarshalContext> ThreadInstance
            = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetThreadInstance() => ThreadInstance.Value;

        public XlMarshalContext()
        {
            DoubleValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(double)));
            DecimalValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(decimal)));
            StringValue = Marshal.AllocCoTaskMem(XLOPER12.MaxStringLength + 1);
            LongValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(long)));

            DoubleToIntPtr(0);
            LongToIntPtr(0);
            DecimalToIntPtr(0);
            StringToIntPtr("");
        }

        ~XlMarshalContext()
        {
            Marshal.FreeCoTaskMem(DoubleValue);
            Marshal.FreeCoTaskMem(DecimalValue);
            Marshal.FreeCoTaskMem(StringValue);
            Marshal.FreeCoTaskMem(LongValue);
        }
    }
}
