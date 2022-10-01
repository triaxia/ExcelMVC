using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
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
