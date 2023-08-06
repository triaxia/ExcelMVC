using System;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        // See https://aka.ms/new-console-template for more information
        Console.WriteLine("Hello, World!");

        var template = "expudf(idx)";
        var content = string.Join(Environment.NewLine, Enumerable.Range(0, 5000)
            .Select(x => template.Replace("idx", $"{x}")));
        Console.ReadKey();
    }
}

