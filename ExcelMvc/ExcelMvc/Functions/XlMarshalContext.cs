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
        private readonly IntPtr ObjectValue;

        // thread affinity for return pointers...
        private readonly static ThreadLocal<XlMarshalContext> ThreadInstance
            = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetThreadInstance() => ThreadInstance.Value;

        public XlMarshalContext()
        {
            DoubleValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(double)));
            IntValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));
            ShortValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(short)));
            ObjectValue = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(XLOPER12)));

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
            if (ObjectValue != IntPtr.Zero) 
            {
                XLOPER12 *p = (XLOPER12 *) ObjectValue.ToPointer();
                p->Dispose();
                Marshal.FreeCoTaskMem(ObjectValue);
            }
        }
    }
}
