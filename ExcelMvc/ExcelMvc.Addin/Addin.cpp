/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Developer:         Wolfgang Stamm, Germany

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
#include "stdafx.h"
#include "ClrRuntimeHost.h"

/*
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
*/

const int NumberOfParameters = 11;
static LPWSTR rgFuncs[][NumberOfParameters] =
{
	{ L"ExcelMvcRunCommandAction", L"I", L"ExcelMvcRunCommandAction", L"", L"0", L"ExcelMvc", NULL, NULL, NULL, NULL, NULL },
	{ L"ExcelMvcAttach", L"I", L"ExcelMvcAttach", L"", L"2", L"ExcelMvc", NULL, NULL, NULL, NULL, NULL },
	{ L"ExcelMvcDetach", L"I", L"ExcelMvcDetach", L"", L"2", L"ExcelMvc", NULL, NULL, NULL, NULL, NULL }
};

BOOL StartAddinClrHost()
{
	TCHAR buffer[256] = { 0 };
	::GetModuleFileName(Constants::Dll, buffer, sizeof(buffer) / sizeof(WCHAR));

	// trim off file name
	int pos = wcslen(buffer);
	while (--pos >= 0 && buffer[pos] != '\\');
	buffer[pos] = 0;
#if CLR2
    static LPCTSTR clrVersion = L"v2.0.50727";
#elif CLR4
	static LPCTSTR clrVersion = L"v4.0.30319";
#endif
    ClrRuntimeHost::Start(clrVersion, L"ExcelMvc", buffer);
	BOOL result = ClrRuntimeHost::TestAndDisplayError();
	if (result)
	{
		// create a scratch book to get Excel registered with the ROT
		Excel12f(xlcEcho, 0, 1, (LPXLOPER12) TempBool12(false));
		Excel12f(xlcNew, 0, 1, (LPXLOPER12) TempInt12(5));
		Excel12f(xlcWorkbookInsert, 0, 1, (LPXLOPER12) TempInt12(6));

		ClrRuntimeHost::CallStaticMethod(L"ExcelMvc.Runtime.Interface", L"Attach");

		// remove scratch book
		Excel12f(xlcFileClose, 0, 1, (LPXLOPER12) TempBool12(false));
		Excel12f(xlcEcho, 0, 1, (LPXLOPER12) TempBool12(true));

		result = ClrRuntimeHost::TestAndDisplayError();
	}
	return result;
}

void StopAddinClrHost()
{
	ClrRuntimeHost::Stop();
}

BOOL __stdcall xlAutoOpen(void)
{
	static XLOPER12 xDLL;
	Excel12f(xlGetName, &xDLL, 0);

	int count = sizeof(rgFuncs) / (sizeof(rgFuncs[0][0]) * NumberOfParameters);
	for (int idx = 0; idx < count; idx++)
	{
		int macroType = wcscmp(rgFuncs[idx][4], L"0") == 0 ? 0 : 1;
		Excel12f
		(
			xlfRegister, 0, 7,
			(LPXLOPER12) &xDLL,
			(LPXLOPER12) TempStr12(rgFuncs[idx][0]),
			(LPXLOPER12) TempStr12(rgFuncs[idx][1]),
			(LPXLOPER12) TempStr12(rgFuncs[idx][2]),
			(LPXLOPER12) TempStr12(rgFuncs[idx][3]),
			(LPXLOPER12) TempInt12(macroType),
			(LPXLOPER12) TempStr12(rgFuncs[idx][5])/*,
			(LPXLOPER12) TempStr12(rgFuncs[idx][6]),
			(LPXLOPER12) TempStr12(rgFuncs[idx][7]),
			(LPXLOPER12) TempStr12(rgFuncs[idx][8]),
			(LPXLOPER12) TempStr12(rgFuncs[idx][9]),
			(LPXLOPER12) TempStr12(rgFuncs[idx][10])*/
		);
	}

	Excel12f(xlFree, 0, 1, (LPXLOPER12) &xDLL);

	return StartAddinClrHost();
}

BOOL __stdcall xlAutoClose(void)
{
	return TRUE;
}

BOOL __stdcall ExcelMvcRunCommandAction(void)
{
	ClrRuntimeHost::CallStaticMethod(L"ExcelMvc.Runtime.Interface", L"FireClicked");
	return ClrRuntimeHost::TestAndDisplayError();
}

BOOL __stdcall ExcelMvcAttach(void)
{
	ClrRuntimeHost::CallStaticMethod(L"ExcelMvc.Runtime.Interface", L"Attach");
	return ClrRuntimeHost::TestAndDisplayError();
}

BOOL __stdcall ExcelMvcDetach(void)
{
	ClrRuntimeHost::CallStaticMethod(L"ExcelMvc.Runtime.Interface", L"Detach");
	return ClrRuntimeHost::TestAndDisplayError();
}

