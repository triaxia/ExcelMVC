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

namespace ExcelMvc.Functions
{
    [Flags]
    public enum XlTypes
    {
        xltypeNum = 0x0001,
        xltypeStr = 0x0002,
        xltypeBool = 0x0004,
        xltypeRef = 0x0008,
        xltypeErr = 0x0010,
        xltypeFlow = 0x0020,
        xltypeMulti = 0x0040,
        xltypeMissing = 0x0080,
        xltypeNil = 0x0100,
        xltypeSRef = 0x0400,
        xltypeInt = 0x0800,
        xlbitXLFree = 0x1000,
        xlbitDLLFree = 0x4000,
        xltypeBigData = xltypeStr | xltypeInt
    }

    public class ExcelMissing
    {
        private ExcelMissing() { }
        public static readonly ExcelMissing Value = new ExcelMissing();
        public override string ToString() => "";
        public static bool IsMe(object value) => value == Value;
    }

    public class ExcelEmpty
    {
        private ExcelEmpty() { }
        public override string ToString() => "";
        public static readonly ExcelEmpty Value = new ExcelEmpty();
        public static bool IsMe(object value) => value == Value;
    }
}

