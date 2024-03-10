using Addin.Interfaces;
using System;
using System.Linq;
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
        [MarshalAs(UnmanagedType.U1)]
        public byte Type;

        public Argument(ArgumentAttribute rhs)
        {
            Name = rhs.Name;
            Description = rhs.Description;
            Type = 1; //TODO
        }
    }

    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void FunctionCallback(IntPtr args);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct Function
    {
        public const ushort MaxArguments = 32;
        [MarshalAs(UnmanagedType.U4)]
        public int Index;
        [MarshalAs(UnmanagedType.U1)]
        public byte ReturnType;
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

        public Function(int index, FunctionAttribute rhs, Argument[] arguments, IntPtr callback)
        {
            Index = index;
            Callback = callback; // Marshal.GetFunctionPointerForDelegate(callback);
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
            ReturnType = 1; // TODO
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
