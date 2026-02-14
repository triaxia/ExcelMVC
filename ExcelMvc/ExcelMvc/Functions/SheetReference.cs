using Function.Interfaces;
using System;

namespace ExcelMvc.Functions
{
    internal class SheetReference
    {
        public RangeReference Range { get; }
        public IntPtr SheetID { get; }
        public SheetReference(RangeReference range, IntPtr sheetID)
        {
            Range = range;
            SheetID = sheetID;
        }
    }
}
