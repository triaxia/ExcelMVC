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
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct Argument
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Name;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Description;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Type;

        public Argument(ParameterInfo parameter, ExcelArgumentAttribute argument)
        {
            if (argument == null)
            {
                Name = parameter.Name;
                Description = "";
            }
            else
            {
                Name = argument.Name;
                Description = argument.Description;
            }
            Type = parameter.ParameterType.FullName;
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct Function
    {
        public const ushort MaxArguments = 32;
        [MarshalAs(UnmanagedType.U4)]
        public int Index;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string ReturnType;
        // ulong works too
        //[MarshalAs(UnmanagedType.U8)]
        //public ulong Callback;
        public IntPtr Callback;
        [MarshalAs(UnmanagedType.U1)]
        public byte FunctionType;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsVolatile;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsMacro;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsAsync;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsThreadSafe;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsClusterSafe;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Category;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Name;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Description;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string HelpTopic;
        [MarshalAs(UnmanagedType.U1)]
        public byte ArgumentCount;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = MaxArguments)]
        public Argument[] Arguments;

        public Function(int index, ExcelFunctionAttribute rhs, Argument[] arguments
            , IntPtr callback, Type returnType)
        {
            Index = index;
            Callback = callback; 
            FunctionType = rhs.FunctionType;
            IsVolatile = rhs.IsVolatile;
            IsMacro = rhs.IsMacro;
            IsAsync = rhs.IsAsync;
            IsThreadSafe = rhs.IsThreadSafe;
            IsClusterSafe = rhs.IsClusterSafe;
            ArgumentCount = (byte)(arguments?.Length ?? 0);
            Category = rhs.Category ?? "";
            Name = rhs.Name ?? "";
            Description = rhs.Description ?? "";
            HelpTopic = rhs.HelpTopic ?? "";
            Arguments = Pad(arguments);
            if (rhs.IsHidden) FunctionType = 0;
            ReturnType = returnType.FullName;
        }

        private static Argument[] Pad(Argument[] arguments)
        {
            var args = (arguments ?? new Argument[] { });
            while (args.Length < MaxArguments)
                args = args.Concat(new[] { new Argument() }).ToArray();
            return args;
        }
    }
}
