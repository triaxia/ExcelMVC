
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;

namespace ExcelMvc.Functions
{
    public static class XLFunctions
    {
        public const int xlSpecial = 0x4000;
        public const int xlCommand = 0x8000;

        public const int xlfRtd = 379;
        public const int xlCaller = 89;
        public const int xlSheetNm = 5 | xlSpecial;
        public const int xlfAddress = 219;
        public const int xlAsyncReturn = 16 | xlSpecial;
        public const int xlFree = 0 | xlSpecial;
        public const int xlGetName = 9 | xlSpecial;
        public const int xlfRegister = 149;
        public const int xlfUnregister = 201;
        public const int xlcEcho = 141 | xlCommand;
        public const int xlcFileClose = 144 | xlCommand;
        public const int xlcWorkbookInsert = 354 | xlCommand;
        public const int xlcNew = 119 | xlCommand;
    }

    public static class XLCall
    {
        private static readonly object ExclusiveGate = new object();

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetModuleHandle(string moduleName);
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        [SuppressUnmanagedCodeSecurity]
        internal unsafe delegate int Excel12vDelegate(int xlfn, int count, XLOPER12** ppopers, XLOPER12* pOperRes);
        internal static Excel12vDelegate Excel12v;

        static XLCall()
        {
            IntPtr hModuleAddress = GetModuleHandle(null);
            IntPtr pfnExcel12v = GetProcAddress(hModuleAddress, "MdCallBack12");
            Excel12v = (Excel12vDelegate)Marshal.GetDelegateForFunctionPointer(pfnExcel12v, typeof(Excel12vDelegate));
        }

        public unsafe static (int status, object result) Call(int xlFunction, params object[] parameters)
        {
            lock (ExclusiveGate)
            {
                var pParameters = (XLOPER12**) IntPtr.Zero;

                var xlOps = new List<XLOPER12>();
                var xlArgs = new List<StructIntPtr<XLOPER12>>();
                parameters = parameters ?? new object[] { };
                try
                {
                    xlArgs = parameters.Select(x=>
                    {
                        var xlOp = new XLOPER12(x);
                        xlOps.Add(xlOp);
                        return new StructIntPtr<XLOPER12>(ref xlOp);
                    }).ToList();

                    pParameters = (XLOPER12**)Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(XLOPER12*)) * parameters.Length);
                    for (int i = 0; i < parameters.Length; i++)
                    {
                        pParameters[i] = (XLOPER12 *)xlArgs[i].Ptr;
                    }

                    var result = new XLOPER12();
                    using (var ptr = new StructIntPtr<XLOPER12>(ref result))
                    {
                        var pResult = (XLOPER12*)ptr.Ptr;
                        var status = Excel12v(xlFunction, parameters.Length, pParameters, pResult);
                        var obj = pResult->ToObject();
                        Excel12v(XLFunctions.xlFree, 1, &pResult, (XLOPER12*)IntPtr.Zero);
                        return (status, obj);
                    }
                }
                finally
                {
                    foreach (var x in xlArgs)
                        x.Dispose();
                    foreach(var x in xlOps)
                        x.Dispose();
                    if ((IntPtr)pParameters != IntPtr.Zero)
                        Marshal.FreeCoTaskMem((IntPtr)pParameters);
                }
            }
        }
    }
}
