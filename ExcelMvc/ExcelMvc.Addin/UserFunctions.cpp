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

extern "C" { extern PFN ExportTable[]; }

struct ExceArgument
{
	LPCWSTR Name;
	LPCWSTR Description;
	LPCWSTR Type;
};

struct ExcelFunction
{
	int Index;
	LPCWSTR ReturnType;
	// "unsigned long long" works too.
	//unsigned long long Callback; 
	void* Callback;
	byte MacroType;
	bool IsVolatile;
	bool IsMacro;
	bool IsAsync;
	bool IsThreadSafe;
	bool IsClusterSafe;
	LPCWSTR Category;
	LPCWSTR Name;
	LPCWSTR Description;
	LPCWSTR HelpTopic;
	byte ArgumentCount;
	ExceArgument Arguments[];
};

const int MAX_ARG_COUNT = 32;
struct FunctionArgs
{
	LPXLOPER12 Result;
	LPXLOPER12 Args[MAX_ARG_COUNT];
	int Index;
};

static std::map<int, LPXLOPER12> FunctionRegIds;
static XLOPER12 xDll;

void RegisterUserFunctions()
{
	Excel12f(xlGetName, &xDll, 0);
}

void UnregisterUserFunctions()
{
	for (auto it = FunctionRegIds.begin(); it != FunctionRegIds.end(); it++)
	{
		Excel12f(xlfUnregister, 0, 1, it->second);
		FreeXLOper12T(it->second);
	}
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDll);
}

void UnregisterUserFunction(int index)
{
	{
		auto it = FunctionRegIds.find(index);
		if (it != FunctionRegIds.end())
		{
			Excel12f(xlfUnregister, 0, 1, it->second);
			FreeXLOper12T(it->second);
			FunctionRegIds.erase(index);
		}
	}
}

LPCWSTR NullCoalesce(LPCWSTR value)
{
	return value == NULL ? L"" : value;
}

LPXLOPER12 TempStr12SpacesPadded(LPCWSTR value, int spaces)
{
	value = NullCoalesce(value);
	auto len = lstrlenW(value);
	auto dest = new WCHAR[len + spaces + 1];
	wmemcpy_s(dest, len, value, len);
	for (auto idx = 0; idx < spaces; idx++)
		dest[len + idx] = ' ';
	dest[len + spaces] = '\0';
	auto result = TempStr12(dest);
	delete[] dest;
	return result;
}

void MakeArgumentList(ExcelFunction* pFunction, std::wstring &names, std::wstring& types)
{
	types = pFunction->IsAsync ? L">" : L"E";
	for (auto idx = 0; idx < pFunction->ArgumentCount; idx++)
	{
		if (idx > 0) names += L",";
		names += NullCoalesce(pFunction->Arguments[idx].Name);
		types += L"E";
	}
	if (pFunction->IsAsync) types += L"X";
	if (pFunction->IsVolatile) types += L"!";
	if (pFunction->IsThreadSafe) types += L"$";
	if (pFunction->IsClusterSafe) types += L"&";
	if (pFunction->IsMacro) types += L"#";
}

void NormaliseHelpTopic(ExcelFunction* pFunction, std::wstring& topic)
{
	topic = NullCoalesce(pFunction->HelpTopic);
	if (topic.find(L"!") != std::wstring::npos)
		return;

	auto lower = topic;
	for (unsigned int idx = 0; idx < lower.size(); idx++)
		lower[idx] = std::tolower(lower[idx]);
	if (lower.find(L"http://") != std::wstring::npos
		|| lower.find(L"https://") != std::wstring::npos)
		topic += L"!0";
}

extern "C" extern ClrRuntimeHost * pClrHost;

 LPXLOPER12 __stdcall RegisterFunction(ExcelFunction* pFunction)
{
	UnregisterUserFunction(pFunction->Index);
	auto regId = new XLOPER12();
	FunctionRegIds[pFunction->Index] = regId;
	ExportTable[pFunction->Index] = (PFN) pFunction->Callback;
	/*
	https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	LPXLOPER12 pxModuleText
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
	LPXLOPER12 pxArgumentHelp245
	*/
	TCHAR pxProcedure[10];
	wsprintf(pxProcedure, L"udf%d", pFunction->Index);

	std::wstring pxArgumentText; std::wstring pxTypeText;
	MakeArgumentList(pFunction, pxArgumentText, pxTypeText);

	auto pxFunctionText = pFunction->Name;

	auto pxMacroType = pFunction->MacroType;

	auto pxCategory = NullCoalesce(pFunction->Category);
	auto pxShortcutText = L"";

	std::wstring pxHelpTopic;
	NormaliseHelpTopic(pFunction, pxHelpTopic);

	auto count = 10 + pFunction->ArgumentCount;
	auto pParams = new LPXLOPER12[count];

	pParams[0] = &xDll;
	pParams[1] = TempStr12(pxProcedure);
	pParams[2] = TempStr12(pxTypeText.c_str());
	pParams[3] = TempStr12(pxFunctionText);
	pParams[4] = TempStr12(pxArgumentText.c_str());
	pParams[5] = TempInt12(pxMacroType);
	pParams[6] = TempStr12(pxCategory);
	pParams[7] = TempStr12(pxShortcutText);
	pParams[8] = TempStr12(pxHelpTopic.c_str());

	// Excel function Wizard truncates total function help text by up to two characters...
	// So we fool it by adding two spaces.
	pParams[9] = pFunction->ArgumentCount == 0 ?
		TempStr12SpacesPadded(pFunction->Description, 2)
		: TempStr12(NullCoalesce(pFunction->Description));

	for (auto idx = 0; idx < pFunction->ArgumentCount; idx++)
	{
		pParams[10 + idx] = idx == pFunction->ArgumentCount - 1 ?
			TempStr12SpacesPadded(pFunction->Arguments[idx].Description, 2)
			:TempStr12(NullCoalesce(pFunction->Arguments[idx].Description));
	}

	Excel12v(xlfRegister, regId, count, pParams);
	delete[] pParams;
	
	/* C# Marshal.DestroyStructure<ExcelFunction> does not delete nested Argument texts, so
	*  delete them here...
	*/
	for (auto idx = 0; idx < pFunction->ArgumentCount; idx++)
	{
		delete[] pFunction->Arguments[idx].Name;
		delete[] pFunction->Arguments[idx].Description;
		pFunction->Arguments[idx].Name = NULL;
		pFunction->Arguments[idx].Description = NULL;
	}

	if (regId != NULL) regId->xltype = regId->xltype | xlbitDLLFree;
	return regId;
 }

 LPXLOPER12 __stdcall AsyncReturn(LPXLOPER12 handle, LPXLOPER12 result)
 {
	 auto status = new XLOPER12();
	 result->xltype = result->xltype | xlbitDLLFree;
	 Excel12(xlAsyncReturn, status, 2, handle, result);
	 status->xltype = status->xltype | xlbitDLLFree;
	 return status;
 }

 LPXLOPER12 __stdcall RtdCall(FunctionArgs* args)
 {
	 auto pParams = new LPXLOPER12[MAX_ARG_COUNT];
	 auto count = 0;
	 auto jdx = 0;
	 for (auto idx = 0; idx < MAX_ARG_COUNT; idx++)
	 {
		 if (args->Args[idx] == NULL || args->Args[idx]->xltype == xltypeNil) continue;
		pParams[jdx++] = args->Args[idx];
		count++;
	 }
	 auto result = new XLOPER12();
	 memset(result, 0, sizeof(XLOPER12));
	 Excel12v(xlfRtd, result, count, pParams);
	 result->xltype = result->xltype | xlbitDLLFree;
	 delete[] pParams;
	 return result;
 }
