using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    [StructLayout(LayoutKind.Sequential)]
    public struct FunctionArgs
    {
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
        public int Index;

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
    public class FunctionArgsBag : IDisposable
    {
        public FunctionArgs ToArgs()
        {
            var args = new FunctionArgs();
            args.Arg00 = p00.Ptr;
            args.Arg01 = p01.Ptr;
            args.Arg02 = p02.Ptr;
            args.Arg03 = p03.Ptr;
            args.Arg04 = p04.Ptr;
            args.Arg05 = p05.Ptr;
            args.Arg06 = p06.Ptr;
            args.Arg07 = p07.Ptr;
            args.Arg08 = p08.Ptr;
            args.Arg09 = p09.Ptr;
            args.Arg10 = p10.Ptr;
            args.Arg11 = p11.Ptr;
            args.Arg12 = p12.Ptr;
            args.Arg13 = p13.Ptr;
            args.Arg14 = p14.Ptr;
            args.Arg15 = p15.Ptr;
            args.Arg16 = p16.Ptr;
            args.Arg17 = p17.Ptr;
            args.Arg18 = p18.Ptr;
            args.Arg19 = p19.Ptr;
            args.Arg20 = p20.Ptr;
            args.Arg21 = p21.Ptr;
            args.Arg22 = p22.Ptr;
            args.Arg23 = p23.Ptr;
            args.Arg24 = p24.Ptr;
            args.Arg25 = p25.Ptr;
            args.Arg26 = p26.Ptr;
            args.Arg27 = p27.Ptr;
            args.Arg28 = p28.Ptr;
            args.Arg29 = p29.Ptr;
            args.Arg30 = p30.Ptr;
            args.Arg31 = p31.Ptr;
            return args;
        }

        public FunctionArgsBag(params object[] args) 
        {
            object getArg(int idx) => args.Length > idx ? args[idx] : null;
            x00 = XLOPER12.FromObject(getArg(0));
            x01 = XLOPER12.FromObject(getArg(1));
            x02 = XLOPER12.FromObject(getArg(2));
            x03 = XLOPER12.FromObject(getArg(3));
            x04 = XLOPER12.FromObject(getArg(4));
            x05 = XLOPER12.FromObject(getArg(5));
            x06 = XLOPER12.FromObject(getArg(6));
            x07 = XLOPER12.FromObject(getArg(7));
            x08 = XLOPER12.FromObject(getArg(8));
            x09 = XLOPER12.FromObject(getArg(9));
            x10 = XLOPER12.FromObject(getArg(10));
            x11 = XLOPER12.FromObject(getArg(11));
            x12 = XLOPER12.FromObject(getArg(12));
            x13 = XLOPER12.FromObject(getArg(13));
            x14 = XLOPER12.FromObject(getArg(14));
            x15 = XLOPER12.FromObject(getArg(15));
            x16 = XLOPER12.FromObject(getArg(16));
            x17 = XLOPER12.FromObject(getArg(17));
            x18 = XLOPER12.FromObject(getArg(18));
            x19 = XLOPER12.FromObject(getArg(19));
            x20 = XLOPER12.FromObject(getArg(20));
            x21 = XLOPER12.FromObject(getArg(21));
            x22 = XLOPER12.FromObject(getArg(22));
            x23 = XLOPER12.FromObject(getArg(23));
            x24 = XLOPER12.FromObject(getArg(24));
            x25 = XLOPER12.FromObject(getArg(25));
            x26 = XLOPER12.FromObject(getArg(26));
            x27 = XLOPER12.FromObject(getArg(27));
            x28 = XLOPER12.FromObject(getArg(28));
            x29 = XLOPER12.FromObject(getArg(29));
            x30 = XLOPER12.FromObject(getArg(30));
            x31 = XLOPER12.FromObject(getArg(31));

            p00 = new StructIntPtr<XLOPER12>(ref x00);
            p01 = new StructIntPtr<XLOPER12>(ref x01);
            p02 = new StructIntPtr<XLOPER12>(ref x02);
            p03 = new StructIntPtr<XLOPER12>(ref x03);
            p04 = new StructIntPtr<XLOPER12>(ref x04);
            p05 = new StructIntPtr<XLOPER12>(ref x05);
            p06 = new StructIntPtr<XLOPER12>(ref x06);
            p07 = new StructIntPtr<XLOPER12>(ref x07);
            p08 = new StructIntPtr<XLOPER12>(ref x08);
            p09 = new StructIntPtr<XLOPER12>(ref x09);
            p10 = new StructIntPtr<XLOPER12>(ref x10);
            p11 = new StructIntPtr<XLOPER12>(ref x11);
            p12 = new StructIntPtr<XLOPER12>(ref x11);
            p13 = new StructIntPtr<XLOPER12>(ref x11);
            p14 = new StructIntPtr<XLOPER12>(ref x11);
            p15 = new StructIntPtr<XLOPER12>(ref x11);
            p16 = new StructIntPtr<XLOPER12>(ref x11);
            p17 = new StructIntPtr<XLOPER12>(ref x11);
            p18 = new StructIntPtr<XLOPER12>(ref x11);
            p19 = new StructIntPtr<XLOPER12>(ref x11);
            p20 = new StructIntPtr<XLOPER12>(ref x20);
            p21 = new StructIntPtr<XLOPER12>(ref x21);
            p22 = new StructIntPtr<XLOPER12>(ref x22);
            p23 = new StructIntPtr<XLOPER12>(ref x23);
            p24 = new StructIntPtr<XLOPER12>(ref x24);
            p25 = new StructIntPtr<XLOPER12>(ref x25);
            p26 = new StructIntPtr<XLOPER12>(ref x26);
            p27 = new StructIntPtr<XLOPER12>(ref x27);
            p28 = new StructIntPtr<XLOPER12>(ref x28);
            p29 = new StructIntPtr<XLOPER12>(ref x29);
            p30 = new StructIntPtr<XLOPER12>(ref x30);
            p31 = new StructIntPtr<XLOPER12>(ref x31);
        }

        public void Dispose()
        {
            using (p00) { };
            using (p01) { };
            using (p02) { };
            using (p03) { };
            using (p04) { };
            using (p05) { };
            using (p06) { };
            using (p07) { };
            using (p08) { };
            using (p09) { };
            using (p10) { };
            using (p11) { };
            using (p12) { };
            using (p13) { };
            using (p14) { };
            using (p15) { };
            using (p16) { };
            using (p17) { };
            using (p18) { };
            using (p19) { };
            using (p20) { };
            using (p21) { };
            using (p22) { };
            using (p23) { };
            using (p24) { };
            using (p25) { };
            using (p26) { };
            using (p27) { };
            using (p28) { };
            using (p29) { };
            using (p30) { };
            using (p31) { };
        }
        private XLOPER12 x00;
        private XLOPER12 x01;
        private XLOPER12 x02;
        private XLOPER12 x03;
        private XLOPER12 x04;
        private XLOPER12 x05;
        private XLOPER12 x06;
        private XLOPER12 x07;
        private XLOPER12 x08;
        private XLOPER12 x09;
        private XLOPER12 x10;
        private XLOPER12 x11;
        private XLOPER12 x12;
        private XLOPER12 x13;
        private XLOPER12 x14;
        private XLOPER12 x15;
        private XLOPER12 x16;
        private XLOPER12 x17;
        private XLOPER12 x18;
        private XLOPER12 x19;
        private XLOPER12 x20;
        private XLOPER12 x21;
        private XLOPER12 x22;
        private XLOPER12 x23;
        private XLOPER12 x24;
        private XLOPER12 x25;
        private XLOPER12 x26;
        private XLOPER12 x27;
        private XLOPER12 x28;
        private XLOPER12 x29;
        private XLOPER12 x30;
        private XLOPER12 x31;
        private StructIntPtr<XLOPER12> p00;
        private StructIntPtr<XLOPER12> p01;
        private StructIntPtr<XLOPER12> p02;
        private StructIntPtr<XLOPER12> p03;
        private StructIntPtr<XLOPER12> p04;
        private StructIntPtr<XLOPER12> p05;
        private StructIntPtr<XLOPER12> p06;
        private StructIntPtr<XLOPER12> p07;
        private StructIntPtr<XLOPER12> p08;
        private StructIntPtr<XLOPER12> p09;
        private StructIntPtr<XLOPER12> p10;
        private StructIntPtr<XLOPER12> p11;
        private StructIntPtr<XLOPER12> p12;
        private StructIntPtr<XLOPER12> p13;
        private StructIntPtr<XLOPER12> p14;
        private StructIntPtr<XLOPER12> p15;
        private StructIntPtr<XLOPER12> p16;
        private StructIntPtr<XLOPER12> p17;
        private StructIntPtr<XLOPER12> p18;
        private StructIntPtr<XLOPER12> p19;
        private StructIntPtr<XLOPER12> p20;
        private StructIntPtr<XLOPER12> p21;
        private StructIntPtr<XLOPER12> p22;
        private StructIntPtr<XLOPER12> p23;
        private StructIntPtr<XLOPER12> p24;
        private StructIntPtr<XLOPER12> p25;
        private StructIntPtr<XLOPER12> p26;
        private StructIntPtr<XLOPER12> p27;
        private StructIntPtr<XLOPER12> p28;
        private StructIntPtr<XLOPER12> p29;
        private StructIntPtr<XLOPER12> p30;
        private StructIntPtr<XLOPER12> p31;
    }
}
