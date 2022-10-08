using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMvc.Functions
{
    public enum XlErrors
    {
        xlerrNull = 0,
        xlerrDiv0 = 7,
        xlerrValue = 15,
        xlerrRef = 23,
        xlerrName = 29,
        xlerrNum = 36,
        xlerrNA = 42,
        xlerrGettingData = 43
    }
}
