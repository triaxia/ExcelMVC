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
static LPCWSTR MvcFunctions[][NumberOfParameters] =
{
	{ L"ExcelMvcAttach", L"I", L"ExcelMvcAttach", L"", L"1", L"ExcelMvc", L"", L"", L"Attach Excel to ExcelMvc", L"", L"" },
	{ L"ExcelMvcDetach", L"I", L"ExcelMvcDetach", L"", L"1", L"ExcelMvc", L"", L"", L"Detach Excel from ExcelMvc", L"", L"" },
	{ L"ExcelMvcShow", L"I", L"ExcelMvcShow", L"", L"1", L"ExcelMvc", L"", L"", L"Shows the ExcelMvc window", L"", L"" },
	{ L"ExcelMvcHide", L"I", L"ExcelMvcHide", L"", L"1", L"ExcelMvc", L"", L"", L"Hides the ExcelMvc window", L"", L"" },
	{ L"ExcelMvcClick", L"I", L"ExcelMvcRunCommandAction", L"", L"2", L"ExcelMvc", L"", L"", L"Called by a command", L"", L"" },
	{ L"ExcelMvcRun", L"I", L"ExcelMvcRun", L"", L"2", L"ExcelMvc", L"", L"", L"Runs the next action in the async queue", L"", L"" }
};

static XLOPER12 RegIds[]
{
	XLOPER12(),
	XLOPER12(),
	XLOPER12(),
	XLOPER12(),
	XLOPER12(),
	XLOPER12()
};

void RegisterMvcFunctions(LPXLOPER12 xdll)
{
	auto count = sizeof(MvcFunctions) / (sizeof(MvcFunctions[0][0]) * NumberOfParameters);
	for (int idx = 0; idx < count; idx++)
	{
		Excel12f
		(
			xlfRegister, &RegIds[idx], NumberOfParameters + 1,
			(LPXLOPER12)xdll,
			(LPXLOPER12)TempStr12(MvcFunctions[idx][0]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][1]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][2]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][3]),
			(LPXLOPER12)TempInt12(_wtoi(MvcFunctions[idx][4])),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][5]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][6]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][7]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][8]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][9]),
			(LPXLOPER12)TempStr12(MvcFunctions[idx][10])
		);
	}
}

void UnregisterMvcFunctions()
{
	auto count = sizeof(RegIds) / sizeof(XLOPER12);
	for (int idx = 0; idx < count; idx++)
	{
		Excel12f(xlfUnregister, 0, 1, &RegIds[idx]);
	}
}


extern "C" __declspec(dllexport)
BOOL __stdcall ExcelMvcAttach(void)
{
	pClrHost->Attach();
	return pClrHost->TestAndDisplayError();
}

extern "C" __declspec(dllexport)
BOOL __stdcall ExcelMvcDetach(void)
{
	pClrHost->Detach();
	return pClrHost->TestAndDisplayError();
}

extern "C" __declspec(dllexport)
BOOL __stdcall ExcelMvcShow(void)
{
	pClrHost->Show();
	return pClrHost->TestAndDisplayError();
}

extern "C" __declspec(dllexport)
BOOL __stdcall ExcelMvcHide(void)
{
	pClrHost->Hide();
	return pClrHost->TestAndDisplayError();
}

extern "C" __declspec(dllexport)
BOOL __stdcall ExcelMvcClick(void)
{
	pClrHost->Click();
	return pClrHost->TestAndDisplayError();
}

extern "C" __declspec(dllexport)
BOOL __stdcall ExcelMvcRun(void)
{
	pClrHost->Run();
	return pClrHost->TestAndDisplayError();
}
