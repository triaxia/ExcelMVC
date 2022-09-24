/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor,
Boston, MA 02110-1301 USA.
*/
#include "pch.h"
#include <XLCALL.H>
#include "framewrk.h"
#include "ClrRuntimeHost.h"

/*

https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
LPXLOPER12 pxProcedure
LPXLOPER12 pxTypeText
LPXLOPER12 pxFunctionText
LPXLOPER12 pxArgumentText
LPXLOPER12 pxMacroType,
LPXLOPER12 pxCategory
LPXLOPER12 pxShortcutText
LPXLOPER12 pxHelpTopic
LPXLOPER12 pxFunctionHelp
LPXLOPER12 pxArgumentHelp1
LPXLOPER12 pxArgumentHelp2
.
LPXLOPER12 pxArgumentHelp255
*/

extern "C"
{
	extern ClrRuntimeHost* pClrHost; 
}

const int NumberOfParameters = 11;
static LPCWSTR UserFunctions[][NumberOfParameters] =
{
	{ L"ExcelMvcUdf", L"QQQQ", L"ExcelMvcAdd", L"", L"1", L"ExcelMvc", L"", L"", L"Add numbers", L"", L"" }
};

static XLOPER12 RegIds[]
{
	XLOPER12(),
};

void RegisterUserFunctions(LPXLOPER12 xdll)
{
	auto count = sizeof(UserFunctions) / (sizeof(UserFunctions[0][0]) * NumberOfParameters);
	for (int idx = 0; idx < count; idx++)
	{
		Excel12f
		(
			xlfRegister, &RegIds[idx], NumberOfParameters + 1,
			(LPXLOPER12)xdll,
			(LPXLOPER12)TempStr12(UserFunctions[idx][0]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][1]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][2]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][3]),
			(LPXLOPER12)TempInt12(_wtoi(UserFunctions[idx][4])),
			(LPXLOPER12)TempStr12(UserFunctions[idx][5]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][6]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][7]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][8]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][9]),
			(LPXLOPER12)TempStr12(UserFunctions[idx][10])
		);
	}
}

void UnregisterUserFunctions()
{
	auto count = sizeof(RegIds) / sizeof(XLOPER12);
	for (int idx = 0; idx < count; idx++)
	{
		Excel12f(xlfUnregister, 0, 1, &RegIds[idx]);
	}
}

extern "C" __declspec(dllexport)
LPXLOPER12 __stdcall ExcelMvcUdf(
	LPXLOPER12 arg1, LPXLOPER12 arg2, LPXLOPER12 arg3, LPXLOPER12 arg4, LPXLOPER12 arg5, LPXLOPER12 arg6,
	LPXLOPER12 arg7, LPXLOPER12 arg8, LPXLOPER12 arg9, LPXLOPER12 arg10, LPXLOPER12 arg11, LPXLOPER12 arg12,
	LPXLOPER12 arg13, LPXLOPER12 arg14, LPXLOPER12 arg15, LPXLOPER12 arg16, LPXLOPER12 arg17, LPXLOPER12 arg18,
	LPXLOPER12 arg19, LPXLOPER12 arg20, LPXLOPER12 arg21, LPXLOPER12 arg22, LPXLOPER12 arg23, LPXLOPER12 arg24,
	LPXLOPER12 arg25, LPXLOPER12 arg26, LPXLOPER12 arg27, LPXLOPER12 arg28, LPXLOPER12 arg29, LPXLOPER12 arg30,
	LPXLOPER12 arg31, LPXLOPER12 arg32)
{
	LPXLOPER12 result = (LPXLOPER12)malloc(sizeof(XLOPER12));
	result->xltype = xltypeInt | xlbitDLLFree;
	void* args[] =
	{
		result,
		arg1,  arg2,  arg3,  arg4,  arg5,  arg6, arg7,  arg8,  arg9,  arg10,
		arg11,  arg12, arg13,  arg14,  arg15,  arg16,  arg17,  arg18, arg19,  arg20,
		arg21,  arg22,  arg23,  arg24,  arg25,  arg26,  arg27,  arg28,  arg29,  arg30,
		arg31,  arg32
	};

	pClrHost->Udf(33, args);
	return result;
}
