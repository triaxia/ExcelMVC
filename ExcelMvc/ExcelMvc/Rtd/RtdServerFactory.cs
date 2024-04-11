/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor,
Boston, MA 02110-1301 USA.
*/

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
        public class RtdComClassFactory : IComClassFactory
        {
            public RtdServer RtdServer;
            public RtdComClassFactory(RtdServer instance)
            {
                RtdServer = instance;
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
                    ppvObject = Marshal.GetIUnknownForObject(RtdServer);
                }
                else if (riid == GuidIRtdServer)
                {
                    ppvObject = Marshal.GetComInterfaceForObject(RtdServer, typeof(IRtdServer));
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
