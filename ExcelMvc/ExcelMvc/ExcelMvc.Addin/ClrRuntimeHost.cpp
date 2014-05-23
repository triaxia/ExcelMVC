/****************************** Module Header ******************************\
* Module Name:  RuntimeHostV4.cpp
* Project:      CppHostCLR
* Copyright (c) Microsoft Corporation.
*
* The code in this file demonstrates using .NET Framework 4.0 Hosting
* Interfaces (http://msdn.microsoft.com/en-us/library/dd380851.aspx) to host
* .NET runtime 4.0, load a .NET assebmly, and invoke a type in the assembly.
*
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/en-us/openness/licenses.aspx#MPL.
* All other rights reserved.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "stdafx.h"
#include <windows.h>
#include <metahost.h>
#include "ClrRuntimeHost.h"

#pragma region Includes and Imports
#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only				\
	high_property_prefixes("_get", "_put", "_putref")		\
	rename("ReportEvent", "InteropServices_ReportEvent")
using namespace mscorlib;
#pragma endregion

WCHAR ClrRuntimeHost::LastError[256] = {};
size_t ClrRuntimeHost::LastErrorCount = sizeof(ClrRuntimeHost::LastError) / sizeof(WCHAR);

static ICLRMetaHost *pMetaHost = NULL;
static ICLRRuntimeInfo *pRuntimeInfo = NULL;

// ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting interfaces
// supported by CLR 4.0. Here we demo the ICorRuntimeHost interface that 
// was provided in .NET v1.x, and is compatible with all .NET Frameworks. 
static ICorRuntimeHost *pCorRuntimeHost = NULL;

// ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting interfaces
// supported by CLR 4.0. Here we demo the ICLRRuntimeHost interface that 
// was provided in .NET v2.0 to support CLR 2.0 new features. 
// ICLRRuntimeHost does not support loading the .NET v1.x runtimes.
static ICLRRuntimeHost *pClrRuntimeHost = NULL;

static IUnknownPtr pAppDomainSetupThunk = NULL;
static IAppDomainSetupPtr pAppDomainSetup = NULL;

static IUnknownPtr pAppDomainThunk = NULL;
static _AppDomainPtr pAppDomain = NULL;

static _AssemblyPtr pAssembly = NULL;

BOOL
ClrRuntimeHost::Start(PCWSTR pszVersion, PCWSTR pszAssemblyName, PCWSTR basePath)
{
	LastError[0] = 0;
	bstr_t bstrAssemblyName(pszAssemblyName);
	bstr_t bstrBasePath(basePath);

	HRESULT hr;
	hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost));
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"CLRCreateInstance failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	// Get the ICLRRuntimeInfo corresponding to a particular CLR version. It 
	// supersedes CorBindToRuntimeEx with STARTUP_LOADER_SAFEMODE.
	hr = pMetaHost->GetRuntime(pszVersion, IID_PPV_ARGS(&pRuntimeInfo));
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"ICLRMetaHost::GetRuntime (%s) failed w/hr 0x%08lx\n", pszVersion, hr);
		goto Cleanup;
	}

	// Check if the specified runtime can be loaded into the process. This 
	// method will take into account other runtimes that may already be 
	// loaded into the process and set pbLoadable to TRUE if this runtime can 
	// be loaded in an in-process side-by-side fashion. 
	BOOL fLoadable;
	hr = pRuntimeInfo->IsLoadable(&fLoadable);
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"ICLRRuntimeInfo::IsLoadable failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	if (!fLoadable)
	{
        swprintf(LastError, LastErrorCount, L".NET runtime %s cannot be loaded\n", pszVersion);
		goto Cleanup;
	}


	// Load the CLR into the current process and return a runtime interface 
	// pointer. ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting  
	// interfaces supported by CLR 4.0. Here we demo the ICorRuntimeHost 
	// interface that was provided in .NET v1.x, and is compatible with all 
	// .NET Frameworks. 
	hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&pCorRuntimeHost));
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"ICLRRuntimeInfo::GetInterface failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}


	/*
	
	// Load the CLR into the current process and return a runtime interface 
	// pointer. ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting  
	// interfaces supported by CLR 4.0. Here we demo the ICorRuntimeHost 
	// interface that was provided in .NET v1.x, and is compatible with all 
	// .NET Frameworks. 
	hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&pClrRuntimeHost));
	if (FAILED(hr))
	{
		swprintf(LastError, L"ICLRRuntimeInfo::GetInterface failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	*/


	// Start the CLR.
	hr = pCorRuntimeHost->Start();
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"CLR failed to start w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pCorRuntimeHost->CreateDomainSetup(&pAppDomainSetupThunk);
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"ICorRuntimeHost::CreateDomainSetup failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}
	hr = pAppDomainSetupThunk->QueryInterface(IID_PPV_ARGS(&pAppDomainSetup));
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"Failed to get AppDomainSetup w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}
	hr = pAppDomainSetup->put_ApplicationBase(bstrBasePath);
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"Failed to AppDomainSetup.ApplicationBase w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	// Get a pointer to the default AppDomain in the CLR.
	//hr = pCorRuntimeHost->GetDefaultDomain(&spAppDomainThunk);
	hr = pCorRuntimeHost->CreateDomainEx(L"ExcelMvc", pAppDomainSetup, NULL, &pAppDomainThunk);
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"ICorRuntimeHost::GetDefaultDomain failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pAppDomainThunk->QueryInterface(IID_PPV_ARGS(&pAppDomain));
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"Failed to get AppDomain w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pAppDomain->Load_2(bstrAssemblyName, &pAssembly);
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"Failed to load the assembly w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	return TRUE;
Cleanup:
	Stop();
	return FALSE;
}

void
ClrRuntimeHost::CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName)
{
	LastError[0] = 0;

	bstr_t bstrClassName(pszClassName);
	bstr_t bstrMethodName(pszMethodName);
	SAFEARRAY *psaMethodArgs = NULL;
	variant_t vtEmpty;
	variant_t vtReturn;

	_TypePtr spType = NULL;
	HRESULT hr = pAssembly->GetType_2(bstrClassName, &spType);
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"Failed to get the type %s w/hr 0x%08lx\n", pszClassName, hr);
		goto Cleanup;
	}

	psaMethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, 0);
	hr = spType->InvokeMember_3(
		bstrMethodName,
		static_cast<BindingFlags>(BindingFlags_InvokeMethod | BindingFlags_Static | BindingFlags_Public),
		NULL,
		vtEmpty, 
		psaMethodArgs,
		&vtReturn);
	if (FAILED(hr))
	{
        swprintf(LastError, LastErrorCount, L"Failed to invoke %s w/hr 0x%08lx\n", pszMethodName, hr);
		goto Cleanup;
	}

	return;

Cleanup:
	if (psaMethodArgs)
	{
		SafeArrayDestroy(psaMethodArgs);
	}
	if (spType)
	{
		spType->Release();
	}
}

void
ClrRuntimeHost::Stop()
{
	if (pMetaHost)
	{
		pMetaHost->Release();
		pMetaHost = NULL;
	}

	if (pRuntimeInfo)
	{
		pRuntimeInfo->Release();
		pRuntimeInfo = NULL;
	}

	if (pCorRuntimeHost)
	{
		pCorRuntimeHost->Stop();
		pCorRuntimeHost->Release();
		pCorRuntimeHost = NULL;
	}

	if (pClrRuntimeHost)
	{
		pClrRuntimeHost->Stop();
		pClrRuntimeHost->Release();
		pClrRuntimeHost = NULL;
	}

	if (pAppDomainSetupThunk)
	{
		pAppDomainSetupThunk->Release();
		pAppDomainSetupThunk = NULL;
	}

	if (pAppDomainSetup)
	{
		pAppDomainSetup->Release();
		pAppDomainSetup = NULL;
	}

	if (pAppDomainThunk)
	{
		pAppDomainThunk->Release();
		pAppDomainThunk = NULL;
	}

	if (pAppDomain)
	{
		pAppDomain->Release();
		pAppDomain = NULL;
	}

	if (pAssembly)
	{
		pAssembly->Release();
		pAssembly = NULL;
	}
}

BOOL
ClrRuntimeHost::TestAndDisplayError()
{
	BOOL result = wcslen(LastError) == 0;
	if (!result)
		MessageBox(0, LastError, L"ExcelMvc", MB_OK + MB_ICONERROR);
	return result;
}


