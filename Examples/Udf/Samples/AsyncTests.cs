﻿using ExcelMvc.Functions;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace Samples
{
    public static class AsyncTests
    {
        [ExcelFunction(Name = "uAsync", IsAsync = true)]
        public static void Async(double arg1, double arg2, IntPtr handle)
        {
            Task.Run(()=>
            {
                Thread.Sleep(2000);
                var sum = arg1 + arg2;
                XlCall.SetAsyncResult(handle, sum);
            });
        }
    }
}