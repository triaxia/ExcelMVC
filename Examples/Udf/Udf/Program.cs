using System;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        // See https://aka.ms/new-console-template for more information
        Console.WriteLine("Hello, World!");

        var template = "extern \"C\" __declspec(dllexport) LPXLOPER12 __stdcall Udf{index}(LPXLOPER12 arg1, LPXLOPER12 arg2, LPXLOPER12 arg3, LPXLOPER12 arg4, LPXLOPER12 arg5, LPXLOPER12 arg6, LPXLOPER12 arg7, LPXLOPER12 arg8, LPXLOPER12 arg9, LPXLOPER12 arg10, LPXLOPER12 arg11, LPXLOPER12 arg12, LPXLOPER12 arg13, LPXLOPER12 arg14, LPXLOPER12 arg15, LPXLOPER12 arg16, LPXLOPER12 arg17, LPXLOPER12 arg18, LPXLOPER12 arg19, LPXLOPER12 arg20, LPXLOPER12 arg21, LPXLOPER12 arg22, LPXLOPER12 arg23, LPXLOPER12 arg24, LPXLOPER12 arg25, LPXLOPER12 arg26, LPXLOPER12 arg27, LPXLOPER12 arg28, LPXLOPER12 arg29, LPXLOPER12 arg30, LPXLOPER12 arg31, LPXLOPER12 arg32) { return Udf32({seq}, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30, arg31, arg32); }";
        var content = string.Join(Environment.NewLine, Enumerable.Range(0, 5000)
            .Select(x => template.Replace("{index}", $"{x:0000}").Replace("{seq}", $"{x}")));
        var names = string.Join(Environment.NewLine, Enumerable.Range(0, 5000).Select(x => $"L\"Udf{x:0000}\","));
        Console.ReadKey();
    }
}

