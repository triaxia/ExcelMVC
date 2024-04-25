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

using System;
using System.Runtime.InteropServices;
using ExcelMvc.Rtd;
namespace ExcelMvc.Functions
{
    using HRESULT = Int32;
    using IID = Guid;
    using CLSID = Guid;

    public static class AddIn
    {
        public static string ModuleFileName { get; private set; }
        public delegate void RegisterFunctionsDelegate(IntPtr functions);
        public static RegisterFunctionsDelegate RegisterFunctions { get; private set; }
        public delegate IntPtr AsyncReturnDelegate(IntPtr handle, IntPtr result);
        public static AsyncReturnDelegate AsyncReturn { get; private set; }
        public delegate IntPtr RtdCallDelegate(IntPtr args);
        public static RtdCallDelegate RtdCall { get; private set; }
        public delegate IntPtr AutoFreeDelegate(IntPtr args);
        public static AutoFreeDelegate AutoFree { get; private set; }

        internal delegate HRESULT FuncDllGetClassObject(CLSID rclsid, IID riid, out IntPtr ppunk);

        [StructLayout(LayoutKind.Sequential)]
        public struct AddInHead
        {
            public IntPtr ModuleFileName;
            public IntPtr pDllGetClassObject;
            public IntPtr pRegisterFunctions;
            public IntPtr pAsyncResult;
            public IntPtr pRtdCall;
            public IntPtr pAutoFree;
        }

        public static void OnAttach(IntPtr head)
        {
            unsafe
            {
                AddInHead* pAddInHead = (AddInHead*)head;
                ModuleFileName = Marshal.PtrToStringAuto(pAddInHead->ModuleFileName);
                FuncDllGetClassObject fnDllGetClassObject = RtdServerFactory.DllGetClassObject;
                GCHandle.Alloc(fnDllGetClassObject);
                pAddInHead->pDllGetClassObject = Marshal.GetFunctionPointerForDelegate(fnDllGetClassObject);
                RegisterFunctions = Marshal.GetDelegateForFunctionPointer<RegisterFunctionsDelegate>(pAddInHead->pRegisterFunctions);
                AsyncReturn = Marshal.GetDelegateForFunctionPointer<AsyncReturnDelegate>(pAddInHead->pAsyncResult);
                RtdCall = Marshal.GetDelegateForFunctionPointer<RtdCallDelegate>(pAddInHead->pRtdCall);
                AutoFree = Marshal.GetDelegateForFunctionPointer<AutoFreeDelegate>(pAddInHead->pAutoFree);
            }
        }
    }
}
