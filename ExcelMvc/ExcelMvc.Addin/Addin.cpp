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
#include "ClrRuntimeHostFactory.h"
#include <new>

extern "C" { extern WCHAR ModuleFileName[]; }

extern void __stdcall RegisterFunctions(void* handle);

typedef HRESULT(__stdcall* PFN_DllGetClassObject)(CLSID clsid, IID iid, LPVOID* ppv);
typedef void(__stdcall* PFN_RegisterFunctions)(void* handle);
typedef void(__stdcall* PFN_AutoOpen)();
typedef void(__stdcall* PFN_AutoClose)();
typedef void(__stdcall* PFN_CalculationCancelled)();
typedef void(__stdcall* PFN_CalculationEnded)();


struct AddInHead
{
	LPWSTR ModuleFileName;
	PFN_RegisterFunctions pRegisterFunctions;
	PFN_DllGetClassObject pDllGetClassObject;
	PFN_AutoOpen pAutoOpen;
	PFN_AutoClose pAutoClose;
	PFN_CalculationCancelled pCalculationCancelled;
	PFN_CalculationEnded pCalculationEnded;
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
	pAddInHead->ModuleFileName = new WCHAR[MAX_PATH];
	memcpy(pAddInHead->ModuleFileName, ModuleFileName, sizeof(WCHAR) * MAX_PATH);
	pAddInHead->pRegisterFunctions = RegisterFunctions;
	pAddInHead->pDllGetClassObject = NULL;
	pAddInHead->pAutoOpen = NULL;
	pAddInHead->pCalculationCancelled = NULL;
	pAddInHead->pCalculationEnded = NULL;
	return pAddInHead;
}

extern "C"
{
	ClrRuntimeHost* pClrHost = nullptr;
}

void StopAddInClrHost()
{
	if (pClrHost == nullptr)
		return;

	pClrHost->Stop();
	delete pClrHost;
	pClrHost = nullptr;
}

BOOL StartAddInClrHost()
{
	StopAddInClrHost();
	pClrHost = ClrRuntimeHostFactory::Create();
	pClrHost->Start(L"ExcelMvc", L"ExcelMvc.Runtime.Interface");
	BOOL result = pClrHost->TestAndDisplayError();
	if (result)
	{
		// attach to ExcelMVC
		pClrHost->Attach(CreateAddInHead());
		result = pClrHost->TestAndDisplayError();
	}
	return result;
}

BOOL __stdcall xlAutoOpen(void)
{
	if (StartAddInClrHost())
		pAddInHead->pAutoOpen();
	return TRUE;
}

BOOL __stdcall xlAutoClose(void)
{
	return TRUE;
}

BOOL __stdcall xlAutoRemove(void)
{
	pAddInHead->pAutoClose();
	StopAddInClrHost();
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

void __stdcall CalculationCancelled()
{
	pAddInHead->pCalculationCancelled();
}

void __stdcall CalculationEnded()
{
	pAddInHead->pCalculationEnded();
}