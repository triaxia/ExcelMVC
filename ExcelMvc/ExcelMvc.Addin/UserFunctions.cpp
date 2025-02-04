/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia

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
#include "CallStatus.h"

extern "C" { extern PFN ExportTable[]; }

struct ExcelArgument
{
	LPCWSTR Name;
	LPCWSTR Description;
	LPCWSTR Type;
};

struct ExcelFunction
{
	LPCWSTR ReturnType;
	// "unsigned long long" works too.
	//unsigned long long Callback; 
	void* Callback;
	bool IsVolatile;
	bool IsMacroType;
	bool IsHidden;
	bool IsAsync;
	bool IsThreadSafe;
	bool IsClusterSafe;
	LPCWSTR Category;
	LPCWSTR Name;
	LPCWSTR Description;
	LPCWSTR HelpTopic;
	byte ArgumentCount;
	ExcelArgument Arguments[64];
};

struct ExcelFunctions
{
	int FunctionCount;
	ExcelFunction Functions[];
};

struct FunctionArgument
{
	LPCWSTR Name;
	LPCWSTR Value;
	void* Any;
};

struct FunctionArguments
{
	int Function;
	byte ArgumentCount;
	FunctionArgument Arguments[];
};

static std::map<int, LPXLOPER12> FunctionRegIds;
void UnregisterFunctions()
{
	for (auto it = FunctionRegIds.begin(); it != FunctionRegIds.end(); it++)
	{
		Excel12f(xlfUnregister, 0, 1, it->second);
		Excel12f(xlFree, 0, 1, it->second);
		delete it->second;
	}
	FunctionRegIds.clear();
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

std::wstring MakeTypeString(LPCWSTR type, LPCWSTR argName)
{
	/*
	// optional parameter to use Q, so we can set default values!
	if (argName[0] == L'[' && argName[lstrlenW(argName) - 1] == L']')
		return L"Q";
	*/

	if (wcscmp(type, L"System.Double") == 0
		|| wcscmp(type, L"System.Float") == 0
		|| wcscmp(type, L"System.UInt32") == 0
		|| wcscmp(type, L"System.DateTime") == 0)
		return L"E";
	if (wcscmp(type, L"System.Boolean") == 0)
		return L"L";
	if (wcscmp(type, L"System.Int16") == 0
		|| wcscmp(type, L"System.Byte") == 0
		|| wcscmp(type, L"System.SByte") == 0)
		return L"M";
	if (wcscmp(type, L"System.Int32") == 0
		|| wcscmp(type, L"System.UInt16") == 0)
		return L"N";
	if (wcscmp(type, L"System.String") == 0)
		return L"C%";
	if (wcscmp(type, L"System.Double[,]") == 0
		|| wcscmp(type, L"System.Double[]") == 0
		|| wcscmp(type, L"System.DateTime[,]") == 0
		|| wcscmp(type, L"System.DateTime[]") == 0
		|| wcscmp(type, L"System.Int32[]") == 0
		|| wcscmp(type, L"System.Int32[,]") == 0)
		return L"K%"; // O% does not work!?
	if (wcscmp(type, L"System.IntPtr") == 0)
		return L"X";
	return L"Q";
}

void MakeArgumentList(ExcelFunction* pFunction, std::wstring& names, std::wstring& types)
{
	types = pFunction->IsAsync ? L">" : MakeTypeString(pFunction->ReturnType, L"");
	for (auto idx = 0; idx < pFunction->ArgumentCount; idx++)
	{
		if (idx > 0) names += L",";
		names += NullCoalesce(pFunction->Arguments[idx].Name);
		types += MakeTypeString(pFunction->Arguments[idx].Type, pFunction->Arguments[idx].Name);
	}
	if (pFunction->IsVolatile) types += L"!";
	if (pFunction->IsThreadSafe && !pFunction->IsMacroType) types += L"$";
	if (pFunction->IsClusterSafe && !pFunction->IsAsync) types += L"&";
	if (pFunction->IsMacroType) types += L"#";
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

void RegisterFunction(ExcelFunction* pFunction, int index, LPXLOPER12 xll)
{
	auto regId = new XLOPER12();
	FunctionRegIds[index] = regId;
	ExportTable[index] = (PFN)pFunction->Callback;
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
	wsprintf(pxProcedure, L"udf%d", index);

	std::wstring pxArgumentText; std::wstring pxTypeText;
	MakeArgumentList(pFunction, pxArgumentText, pxTypeText);

	auto pxFunctionText = pFunction->Name;

	auto pxMacroType = pFunction->IsHidden ? 0 : 1;

	auto pxCategory = NullCoalesce(pFunction->Category);
	auto pxShortcutText = L"";

	std::wstring pxHelpTopic;
	NormaliseHelpTopic(pFunction, pxHelpTopic);

	auto count = 10 + pFunction->ArgumentCount;
	auto parameters = new LPXLOPER12[count];

	parameters[0] = xll;
	parameters[1] = TempStr12(pxProcedure);
	parameters[2] = TempStr12(pxTypeText.c_str());
	parameters[3] = TempStr12(pxFunctionText);
	parameters[4] = TempStr12(pxArgumentText.c_str());
	parameters[5] = TempInt12(pxMacroType);
	parameters[6] = TempStr12(pxCategory);
	parameters[7] = TempStr12(pxShortcutText);
	parameters[8] = TempStr12(pxHelpTopic.c_str());

	// Excel function Wizard truncates total function help text by up to two characters...
	// So we fool it by adding two spaces.
	parameters[9] = pFunction->ArgumentCount == 0 ?
		TempStr12SpacesPadded(pFunction->Description, 2)
		: TempStr12(NullCoalesce(pFunction->Description));

	for (auto idx = 0; idx < pFunction->ArgumentCount; idx++)
	{
		parameters[10 + idx] = idx == pFunction->ArgumentCount - 1 ?
			TempStr12SpacesPadded(pFunction->Arguments[idx].Description, 2)
			: TempStr12(NullCoalesce(pFunction->Arguments[idx].Description));
	}

	Excel12v(xlfRegister, regId, count, parameters);
	FreeAllTempMemory();
	delete[] parameters;
	regId->xltype = regId->xltype | xlbitDLLFree;
}

void __stdcall RegisterFunctions(void* handle)
{
	UnregisterFunctions();

	static XLOPER12 xDll;
	Excel12f(xlGetName, &xDll, 0);

	auto pFunctions = (ExcelFunctions*)handle;
	for (auto index = 0; index < pFunctions->FunctionCount; index++)
	{
		auto x = pFunctions->Functions[index];
		RegisterFunction(&x, index, &xDll);
	}
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDll);
}

LPCALLSTATUS __stdcall SetAsyncValue(LPXLOPER12 handle, LPXLOPER12 value)
{
	auto result = new XLOPER12();
	auto code = Excel12f(xlAsyncReturn, result, 2, handle, value);

	auto cr = new CallStatus();
	cr->Result = result;
	cr->status = code;
	return cr;
}

LPCALLSTATUS __stdcall CallRtd(void* handle)
{
	auto args = (FunctionArguments*)handle;
	auto parameters = new LPXLOPER12[args->ArgumentCount];
	for (auto idx = 0; idx < args->ArgumentCount; idx++)
		parameters[idx] = TempStr12(args->Arguments[idx].Value);

	auto result = new XLOPER12();
	memset(result, 0, sizeof(XLOPER12));
	auto code = Excel12v(xlfRtd, result, args->ArgumentCount, parameters);

	FreeAllTempMemory();
	delete[] parameters;

	auto cr = new CallStatus();
	cr->Result = result;
	cr->status = code;
	return cr;
}

LPCALLSTATUS __stdcall CallAny(void* handle)
{
	auto args = (FunctionArguments*)handle;
	auto parameters = new LPXLOPER12[args->ArgumentCount];
	for (auto idx = 0; idx < args->ArgumentCount; idx++)
		parameters[idx] = (LPXLOPER12) args->Arguments[idx].Any;

	auto result = new XLOPER12();
	memset(result, 0, sizeof(XLOPER12));
	auto code = Excel12v(args->Function, result, args->ArgumentCount, parameters);

	FreeAllTempMemory();
	delete[] parameters;

	auto cr = new CallStatus();
	cr->Result = result;
	cr->status = code;
	return cr;
}