using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Rtd
{
    [Guid("F80F202A-B862-4D50-AA51-F0481781CB4F")]
    [ComVisible(true)]
    [ProgId("ExcelMvc.TestServer")]
    public class TestServer
    {
        public int Add(int a, int b) => a + b;
    }

}
