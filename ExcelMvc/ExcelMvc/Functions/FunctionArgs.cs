using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Sequential)]
    public struct FunctionArgs
    {
        public uint Index;
        public IntPtr Result;
        public IntPtr Arg00;
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

        public IntPtr[] GetArgs(int size = 32) => new IntPtr[]
        {
            Arg00, Arg01, Arg02, Arg03, Arg04, Arg05, Arg06, Arg07, Arg08, Arg09,
            Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19,
            Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29,
            Arg30, Arg31
        }.Take(size).ToArray();

        public string Print(int size)
        {
            string ToString(IntPtr ptr) => ptr == IntPtr.Zero
                ? null : Marshal.PtrToStructure<XLOPER12>(ptr).xltype.ToString();
            return string.Join(System.Environment.NewLine,
                 GetArgs(size).Select(x => $"{ToString(x)}"));
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct FunctionResult
    {
        public uint Index;
        public IntPtr Value;
    }
}
