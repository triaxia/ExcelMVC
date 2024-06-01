using Function.Interfaces;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace Samples
{
    public static class AsyncTests
    {
        private static readonly Semaphore OneWay = new Semaphore(1, 1);
        private static IntPtr Handle;

        [Function(Name = "uAsync", IsAsync = true)]
        public static void Async(double arg1, double arg2, IntPtr handle)
        {
            Handle = Host.Call.GetAsyncHandle(handle);
            Host.Call.SetAsyncResult(Handle, "...");
            if (!OneWay.WaitOne(0)) return;
            Task.Run(()=>
            {
                try
                {
                    Thread.Sleep(2000);
                    Add(arg1, arg2);
                }
                finally
                {
                    OneWay.Release();
                }
            });
        }

        private static void Add(double arg1, double arg2) 
        {
            var sum = arg1 + arg2;
            Host.Call.SetAsyncResult(Handle, sum);
        }
    }
}
