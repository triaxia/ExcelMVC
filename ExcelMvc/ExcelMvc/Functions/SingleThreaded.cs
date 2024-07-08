using System;
using System.Threading;

namespace ExcelMvc.Functions
{
    internal class SingleThreaded : IDisposable
    {
        private readonly SemaphoreSlim Gate = new SemaphoreSlim(1);

        public SingleThreaded()
        {
            Gate.Wait();
        }
        public void Dispose()
        {
            Gate.Release();
        }
    }
}
