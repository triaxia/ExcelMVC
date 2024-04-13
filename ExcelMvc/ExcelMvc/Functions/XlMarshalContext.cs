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
using System.Threading;

namespace ExcelMvc.Functions
{
    public unsafe partial class XlMarshalContext
    {
        private readonly IntPtr DoubleValue;
        private IntPtr StringValue = IntPtr.Zero;
        private IntPtr LargeStringValue = IntPtr.Zero;
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
            DoubleValue = Marshal.AllocCoTaskMem(sizeof(double));
            IntValue = Marshal.AllocCoTaskMem(sizeof(int));
            ShortValue = Marshal.AllocCoTaskMem(sizeof(short));
            ObjectValue = Marshal.AllocCoTaskMem(sizeof(XLOPER12));

            DoubleToIntPtr(0);
            Int32ToIntPtr(0);
            Int16ToIntPtr(0);
            InitObjectValue();
        }

        ~XlMarshalContext()
        {
            Marshal.FreeCoTaskMem(DoubleValue);
            Marshal.FreeCoTaskMem(StringValue);
            Marshal.FreeCoTaskMem(LargeStringValue);
            Marshal.FreeCoTaskMem(IntValue);
            Marshal.FreeCoTaskMem(ShortValue);
            Marshal.FreeCoTaskMem(DoubleArrayValue);
            if (ObjectValue != IntPtr.Zero) 
            {
                FreeObjectValue();
                Marshal.FreeCoTaskMem(ObjectValue);
            }
        }
    }
}
