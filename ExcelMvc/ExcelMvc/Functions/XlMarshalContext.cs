using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        private readonly IntPtr DoubleValue;
        private readonly IntPtr StringValue;

        // thread affinity for return pointers...
        private readonly static ThreadLocal<XlMarshalContext> ThreadInstance
            = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetThreadInstance() => ThreadInstance.Value;

        public XlMarshalContext()
        {
            var size = Marshal.SizeOf(typeof(double));
            DoubleValue = Marshal.AllocCoTaskMem(size);
            StringValue = Marshal.AllocCoTaskMem(XLOPER12.MaxStringLength + 1);
        }

        ~XlMarshalContext()
        {
            Marshal.FreeCoTaskMem(DoubleValue);
            Marshal.FreeCoTaskMem(StringValue);
        }
    }
}
