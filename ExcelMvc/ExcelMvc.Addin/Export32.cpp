#include "pch.h"

#ifndef _M_X64

extern "C" { extern PFN ExportTable[]; }

#define udf(i) extern "C" __declspec(dllexport,naked) void f##i(void){	__asm jmp ExportTable + i * 4 }

udf(0)
udf(1)
udf(2)
udf(3)
udf(4)
udf(5)
udf(6)
udf(7)
udf(8)

#endif