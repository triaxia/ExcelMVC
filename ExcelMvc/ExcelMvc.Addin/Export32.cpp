#include "pch.h"

#ifndef _M_X64

extern "C" { extern PFN ExportTable[]; }

#define udf(i) extern "C" __declspec(dllexport,naked) void f##i(void){	__asm jmp ExportTable + i * 4 }

udf(0)

#endif