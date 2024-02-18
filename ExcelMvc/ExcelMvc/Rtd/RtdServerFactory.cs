using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Rtd
{
    using HRESULT = Int32;
    using IID = Guid;
    using CLSID = Guid;

    public static class RtdServerFactory
    {
        public const HRESULT S_OK = 0;
        public const HRESULT S_FALSE = 1;
        public const HRESULT CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
        public const HRESULT CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
        public const HRESULT E_ACCESSDENIED = unchecked((int)0x80070005);
        public const HRESULT E_INVALIDARG = unchecked((int)0x80070057);
        public const HRESULT E_NOINTERFACE = unchecked((int)0x80004002);
        public const HRESULT E_UNEXPECTED = unchecked((int)0x8000FFFF);

        public const string GuidStringIUnknown = "00000000-0000-0000-C000-000000000046";
        public const string GuidStringClassFactory = "00000001-0000-0000-C000-000000000046";
        public const string GuidStringIRtdServer = "EC0E6191-DB51-11D3-8F3E-00C04F3651B8";
        public static readonly Guid GuidIUnknown = new Guid(GuidStringIUnknown);
        public static readonly Guid GuidClassFactory = new Guid(GuidStringClassFactory);
        public static readonly Guid GuidIRtdServer = new Guid(GuidStringIRtdServer);

        [ComImport]
        [Guid(GuidStringClassFactory)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IComClassFactory
        {
            [PreserveSig]
            HRESULT CreateInstance([In] IntPtr pUnkOuter, [In] ref IID riid, [Out] out IntPtr ppvObject);

            [PreserveSig]
            HRESULT LockServer([In, MarshalAs(UnmanagedType.VariantBool)] bool fLock);
        }

        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.None)]
        public class ComObjectFactory : IComClassFactory
        {
            private object _instance ;
            public ComObjectFactory(object instance)
            {
                _instance = instance;
            }

            public int CreateInstance([In] IntPtr pUnkOuter, [In] ref CLSID riid, [Out] out IntPtr ppvObject)
            {
                ppvObject = IntPtr.Zero;
                if (pUnkOuter != IntPtr.Zero)
                {
                    return CLASS_E_NOAGGREGATION;
                }
                if (riid == GuidIUnknown)
                {
                    ppvObject = Marshal.GetIUnknownForObject(_instance);
                }
                else if (riid == GuidIRtdServer)
                {
                    ppvObject = Marshal.GetComInterfaceForObject(_instance, typeof(IRtdServer));
                }
                else
                {
                    return E_NOINTERFACE;
                }
                return S_OK;
            }

            public int LockServer([In, MarshalAs(UnmanagedType.VariantBool)] bool fLock)
            {
                return S_OK; 
            }
        }

        public static HRESULT DllGetClassObject(CLSID clsid, IID iid, out IntPtr ppunk)
        {
            if (iid != GuidClassFactory)
            {
                ppunk = IntPtr.Zero;
                return E_INVALIDARG;
            }

            var factory = RtdRegistry.FindFactory(clsid);
            if (factory == null)
            {
                ppunk = IntPtr.Zero;
                return CLASS_E_CLASSNOTAVAILABLE;
            }

            IntPtr punkFactory = Marshal.GetIUnknownForObject(factory);
            var hrQI = Marshal.QueryInterface(punkFactory, ref iid, out ppunk);
            Marshal.Release(punkFactory);
            return hrQI == S_OK ? S_OK : E_UNEXPECTED;
        }
    }
}
