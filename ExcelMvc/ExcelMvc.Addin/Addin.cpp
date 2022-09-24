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
#include "ClrRuntimeHostFactory.h"

extern "C" const GUID __declspec(selectany) DIID__Workbook =
{ 0x000208da, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

BOOL IsExcelThere()
{
	IRunningObjectTable* pRot = NULL;
	::GetRunningObjectTable(0, &pRot);

	IEnumMoniker* pEnum = NULL;
	pRot->EnumRunning(&pEnum);
	IMoniker* pMon[1] = { NULL };
	ULONG fetched = 0;
	BOOL found = FALSE;
	while (pEnum->Next(1, pMon, &fetched) == 0)
	{
		IUnknown* pUnknown;
		pRot->GetObject(pMon[0], &pUnknown);
		IUnknown* pWorkbook;
		if (SUCCEEDED(pUnknown->QueryInterface(DIID__Workbook, (void**)&pWorkbook)))
		{
			found = TRUE;
			pWorkbook->Release();
			break;
		}
		pUnknown->Release();
	}

	if (pRot != NULL)
		pRot->Release();

	if (pEnum != NULL)
		pEnum->Release();

	return found;
}

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

const int NumberOfParameters = 11;
static LPCWSTR MvcFuncs[][NumberOfParameters] =
{
	{ L"ExcelMvcClick", L"I", L"ExcelMvcClick", L"", L"2", L"ExcelMvc", L"", L"", L"Called by a command", L"", L"" },
	{ L"ExcelMvcRun", L"I", L"ExcelMvcRun", L"", L"2", L"ExcelMvc", L"", L"", L"Runs the next action in the async queue", L"", L"" },
	{ L"ExcelMvcAttach", L"I", L"ExcelMvcAttach", L"", L"1", L"ExcelMvc", L"", L"", L"Attach Excel to ExcelMvc", L"", L"" },
	{ L"ExcelMvcDetach", L"I", L"ExcelMvcDetach", L"", L"1", L"ExcelMvc", L"", L"", L"Detach Excel from ExcelMvc", L"", L"" },
	{ L"ExcelMvcShow", L"I", L"ExcelMvcShow", L"", L"1", L"ExcelMvc", L"", L"", L"Shows the ExcelMvc window", L"", L"" },
	{ L"ExcelMvcHide", L"I", L"ExcelMvcHide", L"", L"1", L"ExcelMvc", L"", L"", L"Hides the ExcelMvc window", L"", L"" },
	{ L"ExcelMvcUdf", L"QQQQ", L"ExcelMvcAdd", L"", L"1", L"ExcelMvc", L"", L"", L"Add numbers", L"", L"" }
};

static LPCWSTR MethodNames[] =
{
	L"Attach",
	L"Detach",
	L"Show",
	L"Hide",
	L"Run",
	L"Click",
	L"Udf"
};

static XLOPER12 RegIds[]
{
	XLOPER12(),
	XLOPER12(),
	XLOPER12(),
	XLOPER12(),
	XLOPER12(),
	XLOPER12(),
};

void RegisterFunctions(LPXLOPER12 xdll, LPCWSTR funcs[][NumberOfParameters], int count)
{
	for (int idx = 0; idx < count; idx++)
	{
		Excel12f
		(
			xlfRegister, &RegIds[idx], 12,
			(LPXLOPER12)xdll,
			(LPXLOPER12)TempStr12(funcs[idx][0]),
			(LPXLOPER12)TempStr12(funcs[idx][1]),
			(LPXLOPER12)TempStr12(funcs[idx][2]),
			(LPXLOPER12)TempStr12(funcs[idx][3]),
			(LPXLOPER12)TempInt12(_wtoi(funcs[idx][4])),
			(LPXLOPER12)TempStr12(funcs[idx][5]),
			(LPXLOPER12)TempStr12(funcs[idx][6]),
			(LPXLOPER12)TempStr12(funcs[idx][7]),
			(LPXLOPER12)TempStr12(funcs[idx][8]),
			(LPXLOPER12)TempStr12(funcs[idx][9]),
			(LPXLOPER12)TempStr12(funcs[idx][10])
		);
	}
}

void UnregisterFunctions(LPCWSTR funcs[][NumberOfParameters], int count)
{
	for (int idx = 0; idx < count; idx++)
	{
		Excel12f(xlfUnregister, 0, 1, &RegIds[idx]);
	}
}

ClrRuntimeHost* pClrHost = nullptr;

BOOL StartAddinClrHost()
{
	delete pClrHost;
	pClrHost = ClrRuntimeHostFactory::Create();
	pClrHost->Start(L"ExcelMvc", L"ExcelMvc.Runtime.Interface"
		, sizeof(MethodNames) / sizeof(LPCWSTR), MethodNames);
	BOOL result = pClrHost->TestAndDisplayError();

	if (result)
	{
		// insert a book to get Excel registered with the ROT
		Excel12f(xlcEcho, 0, 1, (LPXLOPER12)TempBool12(false));
		Excel12f(xlcNew, 0, 1, (LPXLOPER12)TempInt12(5));
		Excel12f(xlcWorkbookInsert, 0, 1, (LPXLOPER12)TempInt12(6));

		// attach to ExcelMVC
		pClrHost->Call(0, 0, 0);
		result = pClrHost->TestAndDisplayError();

		// close the book
		Excel12f(xlcFileClose, 0, 1, (LPXLOPER12)TempBool12(false));
		Excel12f(xlcEcho, 0, 1, (LPXLOPER12)TempBool12(true));

	}
	return result;
}

void StopAddinClrHost()
{
	if (pClrHost != nullptr) pClrHost->Stop();
}

static int AutoOpenCount = 0;
BOOL __stdcall xlAutoOpen(void)
{
	if (++AutoOpenCount > 1)
		return TRUE;

	static XLOPER12 xDll;
	Excel12f(xlGetName, &xDll, 0);

	RegisterFunctions(&xDll, MvcFuncs, sizeof(MvcFuncs) / (sizeof(MvcFuncs[0][0]) * NumberOfParameters));

	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDll);

	return StartAddinClrHost();
}

BOOL __stdcall xlAutoClose(void)
{
	UnregisterFunctions(MvcFuncs, sizeof(MvcFuncs) / (sizeof(MvcFuncs[0][0]) * NumberOfParameters));
	return TRUE;
}

BOOL __stdcall ExcelMvcAttach(void)
{
	pClrHost->Call(0, 0, nullptr);
	return pClrHost->TestAndDisplayError();
}

BOOL __stdcall ExcelMvcDetach(void)
{
	pClrHost->Call(1, 0, nullptr);
	return pClrHost->TestAndDisplayError();
}

BOOL __stdcall ExcelMvcShow(void)
{
	pClrHost->Call(2, 0, nullptr);
	return pClrHost->TestAndDisplayError();
}

BOOL __stdcall ExcelMvcHide(void)
{
	pClrHost->Call(3, 0, nullptr);
	return pClrHost->TestAndDisplayError();
}

BOOL __stdcall ExcelMvcRun(void)
{
	pClrHost->Call(4, 0, nullptr);
	return pClrHost->TestAndDisplayError();
}

BOOL __stdcall ExcelMvcClick(void)
{
	pClrHost->Call(5, 0, nullptr);
	return pClrHost->TestAndDisplayError();
}

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

	pClrHost->Call(6, 33, args);
	return result;
}

void __stdcall xlAutoFree12(LPXLOPER12 pxFree)
{
	if ((pxFree->xltype & xlbitDLLFree) != xlbitDLLFree)
		return;
	pxFree->xltype = pxFree->xltype & (~xlbitDLLFree);
	FreeXLOper12T(pxFree);
}
