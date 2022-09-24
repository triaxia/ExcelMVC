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

extern "C"
{
	extern ClrRuntimeHost* pClrHost;
}

LPXLOPER12 __stdcall Udf32(int index,
	LPXLOPER12 arg01, LPXLOPER12 arg02, LPXLOPER12 arg03, LPXLOPER12 arg04, LPXLOPER12 arg05, LPXLOPER12 arg06, LPXLOPER12 arg07, LPXLOPER12 arg08, LPXLOPER12 arg09, LPXLOPER12 arg10,
	LPXLOPER12 arg11, LPXLOPER12 arg12, LPXLOPER12 arg13, LPXLOPER12 arg14, LPXLOPER12 arg15, LPXLOPER12 arg16, LPXLOPER12 arg17, LPXLOPER12 arg18, LPXLOPER12 arg19, LPXLOPER12 arg20,
	LPXLOPER12 arg21, LPXLOPER12 arg22, LPXLOPER12 arg23, LPXLOPER12 arg24, LPXLOPER12 arg25, LPXLOPER12 arg26, LPXLOPER12 arg27, LPXLOPER12 arg28, LPXLOPER12 arg29, LPXLOPER12 arg30,
	LPXLOPER12 arg31, LPXLOPER12 arg32)
{
	LPXLOPER12 result = (LPXLOPER12)malloc(sizeof(XLOPER12));
	if (result != NULL)
	{
		result->xltype = xltypeInt | xlbitDLLFree;
		void* args[] =
		{
			(void *)index, result,
			arg01,  arg02,  arg03,  arg04,  arg05,  arg06, arg07,  arg08,  arg09,  arg10,
			arg11,  arg12, arg13,  arg14,  arg15,  arg16,  arg17,  arg18, arg19,  arg20,
			arg21,  arg22,  arg23,  arg24,  arg25,  arg26,  arg27,  arg28,  arg29,  arg30,
			arg31,  arg32
		};
		pClrHost->Udf(33, args);
	}
	return result;
}

/*
* Generated code:
*/
typedef LPXLOPER12(__stdcall* UDF)(LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12);
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall Udf0001(LPXLOPER12 arg1, LPXLOPER12 arg2, LPXLOPER12 arg3, LPXLOPER12 arg4, LPXLOPER12 arg5, LPXLOPER12 arg6, LPXLOPER12 arg7, LPXLOPER12 arg8, LPXLOPER12 arg9, LPXLOPER12 arg10, LPXLOPER12 arg11, LPXLOPER12 arg12, LPXLOPER12 arg13, LPXLOPER12 arg14, LPXLOPER12 arg15, LPXLOPER12 arg16, LPXLOPER12 arg17, LPXLOPER12 arg18, LPXLOPER12 arg19, LPXLOPER12 arg20, LPXLOPER12 arg21, LPXLOPER12 arg22, LPXLOPER12 arg23, LPXLOPER12 arg24, LPXLOPER12 arg25, LPXLOPER12 arg26, LPXLOPER12 arg27, LPXLOPER12 arg28, LPXLOPER12 arg29, LPXLOPER12 arg30, LPXLOPER12 arg31, LPXLOPER12 arg32) { return Udf32(0001, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30, arg31, arg32); }
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall Udf0002(LPXLOPER12 arg1, LPXLOPER12 arg2, LPXLOPER12 arg3, LPXLOPER12 arg4, LPXLOPER12 arg5, LPXLOPER12 arg6, LPXLOPER12 arg7, LPXLOPER12 arg8, LPXLOPER12 arg9, LPXLOPER12 arg10, LPXLOPER12 arg11, LPXLOPER12 arg12, LPXLOPER12 arg13, LPXLOPER12 arg14, LPXLOPER12 arg15, LPXLOPER12 arg16, LPXLOPER12 arg17, LPXLOPER12 arg18, LPXLOPER12 arg19, LPXLOPER12 arg20, LPXLOPER12 arg21, LPXLOPER12 arg22, LPXLOPER12 arg23, LPXLOPER12 arg24, LPXLOPER12 arg25, LPXLOPER12 arg26, LPXLOPER12 arg27, LPXLOPER12 arg28, LPXLOPER12 arg29, LPXLOPER12 arg30, LPXLOPER12 arg31, LPXLOPER12 arg32) { return Udf32(0002, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30, arg31, arg32); }
extern "C"
{
	LPCWSTR UdfProcedures[] = { L"Udf0001", L"Udf0002" };
}

