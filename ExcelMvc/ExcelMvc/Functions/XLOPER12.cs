using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Sequential)]
    public struct Args
    {
        public IntPtr Result;

        public IntPtr Arg01;
        public IntPtr Arg02;
        public IntPtr Arg03;
        public IntPtr Arg04;
        public IntPtr Arg05;
        public IntPtr Arg06;
        public IntPtr Arg07;
        public IntPtr Arg08;
        public IntPtr Arg09;
        public IntPtr Arg10;

        public IntPtr Arg11;
        public IntPtr Arg12;
        public IntPtr Arg13;
        public IntPtr Arg14;
        public IntPtr Arg15;
        public IntPtr Arg16;
        public IntPtr Arg17;
        public IntPtr Arg18;
        public IntPtr Arg19;
        public IntPtr Arg20;

        public IntPtr Arg21;
        public IntPtr Arg22;
        public IntPtr Arg23;
        public IntPtr Arg24;
        public IntPtr Arg25;
        public IntPtr Arg26;
        public IntPtr Arg27;
        public IntPtr Arg28;
        public IntPtr Arg29;
        public IntPtr Arg30;

        public IntPtr Arg31;
        public IntPtr Arg32;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct XLOPER12
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 28)]
        [FieldOffset(0)] public byte[] data;
        [FieldOffset(24)] public uint xltype;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct XLOPER12_num
    {
        [MarshalAs(UnmanagedType.R8)]
        [FieldOffset(0)] public double num;
        [FieldOffset(24)] public uint xltype;
    }
}
