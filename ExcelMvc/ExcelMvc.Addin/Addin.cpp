/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Contributor:       Wolfgang Stamm, Germany (2013)

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
#include "CallStatus.h"

extern "C" { extern WCHAR ModuleFileName[]; }

extern "C" const GUID __declspec(selectany) DIID__Workbook =
{ 0x000208da, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

extern void __stdcall RegisterFunctions(void* handle); 
extern LPCALLSTATUS __stdcall SetAsyncValue(LPXLOPER12 handle, LPXLOPER12 value);
extern LPCALLSTATUS __stdcall CallRtd(void* args);

extern void RegisterMvcFunctions();
extern void UnregisterMvcFunctions();
extern void RegisterUserFunctions();
extern void UnregisterUserFunctions(bool freeXll);

typedef HRESULT(__stdcall* PFN_DllGetClassObject)(CLSID clsid, IID iid, LPVOID* ppv);
typedef void(__stdcall* PFN_RegisterFunctions)(void* handle);
typedef LPCALLSTATUS(__stdcall* PFN_SetAsyncValue)(LPXLOPER12 handle, LPXLOPER12 result);
typedef LPCALLSTATUS(__stdcall* PFN_CallRtd)(void* args);
typedef void(__stdcall* PFN_FreeCallStatus)(LPCALLSTATUS result);

void __stdcall xlAutoFree12(LPXLOPER12 pxFree)
{
	if ((pxFree->xltype & xlbitDLLFree) == xlbitDLLFree)
		return;

	// XLOPER12 returned to Excel are ALL shared thread memory, no
	// free memory should be done here!
}

void __stdcall FreeCallStatus(LPCALLSTATUS pxFree)
{
	if (pxFree->Result != NULL)
	{
		Excel12f(xlFree, NULL, 1, pxFree->Result);
		delete pxFree->Result;
	}
	delete pxFree;
}

struct AddInHead
{
	LPWSTR ModuleFileName;
	PFN_DllGetClassObject pDllGetClassObject;
	PFN_RegisterFunctions pRegisterFunctions;
	PFN_SetAsyncValue pSetAsyncValue;
	PFN_CallRtd pCallRtd;
	PFN_FreeCallStatus pFreeCallStatus;
};

AddInHead* pAddInHead = NULL;
void DeleteAddInHead()
{
	if (pAddInHead == NULL) return;
	delete pAddInHead->ModuleFileName;
	delete pAddInHead;
}

AddInHead* CreateAddInHead()
{
	DeleteAddInHead();
	pAddInHead = new AddInHead();
	pAddInHead->pDllGetClassObject = NULL;
	pAddInHead->ModuleFileName = new WCHAR[MAX_PATH];
	memcpy(pAddInHead->ModuleFileName, ModuleFileName, sizeof(WCHAR) * MAX_PATH);
	pAddInHead->pRegisterFunctions = RegisterFunctions;
	pAddInHead->pSetAsyncValue = SetAsyncValue;
	pAddInHead->pCallRtd = CallRtd;
	pAddInHead->pFreeCallStatus = FreeCallStatus;
	return pAddInHead;
}

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

extern "C"
{
	ClrRuntimeHost* pClrHost = nullptr;
}

BOOL StartAddInClrHost()
{
	delete pClrHost;
	pClrHost = ClrRuntimeHostFactory::Create();
	pClrHost->Start(L"ExcelMvc", L"ExcelMvc.Runtime.Interface");
	BOOL result = pClrHost->TestAndDisplayError();
	if (result)
	{
		// insert a book to get Excel registered with the ROT
		Excel12f(xlcEcho, 0, 1, (LPXLOPER12)TempBool12(false));
		Excel12f(xlcNew, 0, 1, (LPXLOPER12)TempInt12(5));
		Excel12f(xlcWorkbookInsert, 0, 1, (LPXLOPER12)TempInt12(6));

		// attach to ExcelMVC
		pClrHost->Attach(CreateAddInHead());
		result = pClrHost->TestAndDisplayError();

		// close the book
		Excel12f(xlcFileClose, 0, 1, (LPXLOPER12)TempBool12(false));
		Excel12f(xlcEcho, 0, 1, (LPXLOPER12)TempBool12(true));
	}
	return result;
}

void StopAddInClrHost()
{
	if (pClrHost != nullptr) pClrHost->Stop();
}

bool opened = false;

BOOL __stdcall xlAutoOpen(void)
{
	if (!opened)
	{
		opened = true;
		RegisterMvcFunctions();
		RegisterUserFunctions();
		StartAddInClrHost();
	}
	return TRUE;
}

BOOL __stdcall xlAutoClose(void)
{
	/*
	UnregisterMvcFunctions();
	UnregisterUserFunctions(false);
	*/
	return TRUE;
}

HRESULT __stdcall DllRegisterServer()
{
	HRESULT result = S_OK;
	return result;
}

HRESULT __stdcall DllUnregisterServer()
{
	HRESULT result = S_OK;
	return result;
}

HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID iid, void** ppv)
{
	return pAddInHead->pDllGetClassObject(clsid, iid, ppv);
}

HRESULT __stdcall DllCanUnloadNow()
{
	return S_FALSE;
}