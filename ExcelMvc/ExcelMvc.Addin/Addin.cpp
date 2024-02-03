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

typedef HRESULT(__stdcall* PFN_DLLGETCLASSOBJECT)(CLSID clsid, IID iid, LPVOID* ppv);
struct AddInHead
{
	LPWSTR ModuleFileName;
	PFN_DLLGETCLASSOBJECT pDllGetClassObject;
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
	::GetModuleFileName(Constants::Dll, pAddInHead->ModuleFileName, MAX_PATH);
	return pAddInHead;
}

extern void RegisterMvcFunctions();
extern void UnregisterMvcFunctions();
extern void RegisterUserFunctions();
extern void UnregisterUserFunctions();

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

BOOL __stdcall xlAutoOpen(void)
{
	RegisterMvcFunctions();
	RegisterUserFunctions();
	StartAddInClrHost();
	return TRUE;
}

BOOL __stdcall xlAutoClose(void)
{
	UnregisterMvcFunctions();
	UnregisterUserFunctions();
	return TRUE;
}

void __stdcall xlAutoFree12(LPXLOPER12 pxFree)
{
	if ((pxFree->xltype & xlbitDLLFree) != xlbitDLLFree)
		return;
	pxFree->xltype = pxFree->xltype & (~xlbitDLLFree);
	FreeXLOper12T(pxFree);
	free(pxFree);
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
	HRESULT result = S_OK;
	return result;
}

HRESULT __stdcall DllCanUnloadNow()
{
	return S_FALSE;
}