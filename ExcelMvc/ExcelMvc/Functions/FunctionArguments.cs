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
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct FunctionArgument
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Name;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Value;

        public FunctionArgument(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct FunctionArguments
    {
        public const ushort MaxArguments = 64;
        [MarshalAs(UnmanagedType.U1)]
        public byte ArgumentCount;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = MaxArguments)]
        public FunctionArgument[] Arguments;

        public FunctionArguments(FunctionArgument[] arguments)
        {
            ArgumentCount = (byte)arguments.Length;
            Arguments = Pad(arguments);
        }

        private static FunctionArgument[] Pad(FunctionArgument[] arguments)
        {
            var args = arguments ?? new FunctionArgument[] { };
            var count = MaxArguments - args.Length;
            if (count > 0)
                args = args.Concat(Enumerable.Range(0, count).Select(_ => new FunctionArgument()))
                    .ToArray();
            return args;
        }
    }
}
