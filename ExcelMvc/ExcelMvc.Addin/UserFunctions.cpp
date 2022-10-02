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

struct ExceArgument
{
	LPCWSTR Name;
	LPCWSTR Description;
};

struct ExcelFunction
{
	int Index;
	byte MacroType;
	bool IsVolatile;
	bool IsMacro;
	bool IsAnyc;
	bool IsThreadSafe;
	bool IsClusterSafe;
	LPCWSTR Category;
	LPCWSTR Name;
	LPCWSTR Description;
	LPCWSTR HelpTopic;
	byte ArgumentCount;
	ExceArgument Arguments[];
};

static std::map<int, LPXLOPER12> RegIds;
static XLOPER12 xDll;

void RegisterUserFunctions()
{
	Excel12f(xlGetName, &xDll, 0);
}

void UnregisterUserFunctions()
{
	for (auto it = RegIds.begin(); it != RegIds.end(); it++)
	{
		Excel12f(xlfUnregister, 0, 1, it->second);
		FreeXLOper12T(it->second);
	}
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDll);
}

void UnregisterUserFunction(int index)
{
	auto it = RegIds.find(index);
	if (it != RegIds.end())
	{
		Excel12f(xlfUnregister, 0, 1, it->second);
		FreeXLOper12T(it->second);
		RegIds.erase(index);
	}
}

LPCWSTR NullCoalesce(LPCWSTR value)
{
	return value == NULL ? L"" : value;
}

void MakeArgumentList(ExcelFunction* pFunction, std::wstring &names, std::wstring& types)
{
	for (auto idx = 0; idx < pFunction->ArgumentCount; idx++)
	{
		if (idx > 0) names += L",";
		names += NullCoalesce(pFunction->Arguments[idx].Name);
		types += L"Q";
	}
	types += L"Q";
}

void NormaliseHelpTopic(ExcelFunction* pFunction, std::wstring& topic)
{
	topic = NullCoalesce(pFunction->HelpTopic);
	if (topic.find(L"!") != std::wstring::npos)
		return;

	auto lower = topic;
	for (auto idx = 0; idx < lower.size(); idx++)
		lower[idx] = std::tolower(lower[idx]);
	if (lower.find(L"http://") != std::wstring::npos
		|| lower.find(L"https://") != std::wstring::npos)
		topic += L"!0";
}


extern "C" __declspec(dllexport) LPXLOPER12 __stdcall RegisterFunction(void* ptr)
{
	ExcelFunction* pFunction = (ExcelFunction*)ptr;
	UnregisterUserFunction(pFunction->Index);
	auto regId = (LPXLOPER12)malloc(sizeof(XLOPER12));
	RegIds[pFunction->Index] = regId;

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
	wsprintf(pxProcedure, L"Udf%04d", pFunction->Index);

	std::wstring pxArgumentText; std::wstring pxTypeText;
	MakeArgumentList(pFunction, pxArgumentText, pxTypeText);

	if (pFunction->IsVolatile) pxTypeText += L"!";
	if (pFunction->IsThreadSafe) pxTypeText += L"$";
	if (pFunction->IsClusterSafe) pxTypeText += L"&";
	if (pFunction->IsMacro) pxTypeText += L"#";
	if (pFunction->IsAnyc) pxTypeText = std::wstring(L">") + pxTypeText + L"X";


	auto pxFunctionText = pFunction->Name;

	auto pxMacroType = pFunction->MacroType;

	auto pxCategory = NullCoalesce(pFunction->Category);
	auto pxShortcutText = L"";

	std::wstring pxHelpTopic;
	NormaliseHelpTopic(pFunction, pxHelpTopic);

	auto pxFunctionHelp = NullCoalesce(pFunction->Description);

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
	pParams[9] = TempStr12(pxFunctionHelp);
	for (auto idx = 0; idx < pFunction->ArgumentCount; idx++)
		pParams[10 + idx] = (LPXLOPER12)TempStr12(NullCoalesce(pFunction->Arguments[idx].Description));

	Excel12v(xlfRegister, regId, count, pParams);
	FreeAllTempMemory();

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

	return regId;
}