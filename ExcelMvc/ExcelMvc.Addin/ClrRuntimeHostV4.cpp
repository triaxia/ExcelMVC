#include "pch.h"
#include <metahost.h>
#include "ClrRuntimeHostV4.h"

#pragma region Includes and Imports
#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only				\
	high_property_prefixes("_get", "_put", "_putref")		\
	rename("ReportEvent", "InteropServices_ReportEvent") \
    rename("or", "or_arg")
using namespace mscorlib;
#pragma endregion

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

static _TypePtr pClass = NULL;


BOOL 
ClrRuntimeHostV4::Start(PCWSTR pszVersion, PCWSTR pszAssemblyName,
	PCWSTR pszClassName, int argc, PCWSTR _[])
{
	ClearError();

	auto basePath = GetBasePath();
	bstr_t bstrAssemblyName(pszAssemblyName);
	bstr_t bstrBasePath(basePath.c_str());

	HRESULT hr;
	hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost));
	if (FAILED(hr))
	{
        FormatError(L"CLRCreateInstance failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	// Get the ICLRRuntimeInfo corresponding to a particular CLR version. It 
	// supersedes CorBindToRuntimeEx with STARTUP_LOADER_SAFEMODE.
	hr = pMetaHost->GetRuntime(pszVersion, IID_PPV_ARGS(&pRuntimeInfo));
	if (FAILED(hr))
	{
        FormatError(L"ICLRMetaHost::GetRuntime (%s) failed w/hr 0x%08lx\n", pszVersion, hr);
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
        FormatError(L"ICLRRuntimeInfo::IsLoadable failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	if (!fLoadable)
	{
        FormatError(L".NET runtime %s cannot be loaded\n", pszVersion);
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
        FormatError(L"ICLRRuntimeInfo::GetInterface failed w/hr 0x%08lx\n", hr);
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
        FormatError(L"CLR failed to start w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pCorRuntimeHost->CreateDomainSetup(&pAppDomainSetupThunk);
	if (FAILED(hr))
	{
        FormatError(L"ICorRuntimeHost::CreateDomainSetup failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}
	hr = pAppDomainSetupThunk->QueryInterface(IID_PPV_ARGS(&pAppDomainSetup));
	if (FAILED(hr))
	{
        FormatError(L"Failed to get AppDomainSetup w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}
	hr = pAppDomainSetup->put_ApplicationBase(bstrBasePath);
	if (FAILED(hr))
	{
        FormatError(L"Failed to AppDomainSetup.ApplicationBase w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

    // set app config file is there is one matching *.dll.config in the base path
    TCHAR configFile[MAX_PATH];
    if (FindAppConfig(basePath.c_str(), configFile, MAX_PATH))
    {
        bstr_t bstrconfigFile(configFile);
        hr = pAppDomainSetup->put_ConfigurationFile(bstrconfigFile);
        if (FAILED(hr))
        {
            FormatError(L"Failed to AppDomainSetup.ConfigurationFile w/hr 0x%08lx\n", hr);
            goto Cleanup;
        }
    }

	// Get a pointer to the default AppDomain in the CLR.
	//hr = pCorRuntimeHost->GetDefaultDomain(&spAppDomainThunk);
	hr = pCorRuntimeHost->CreateDomainEx(L"ExcelMvc", pAppDomainSetup, NULL, &pAppDomainThunk);
	if (FAILED(hr))
	{
        FormatError(L"ICorRuntimeHost::GetDefaultDomain failed w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pAppDomainThunk->QueryInterface(IID_PPV_ARGS(&pAppDomain));
	if (FAILED(hr))
	{
        FormatError(L"Failed to get AppDomain \"ExcelMvc\" w/hr 0x%08lx\n", hr);
		goto Cleanup;
	}

	hr = pAppDomain->Load_2(bstrAssemblyName, &pAssembly);
	if (FAILED(hr))
	{
        FormatError(L"Failed to load assembly \"%s\" w/hr 0x%08lx\n", bstrAssemblyName, hr);
		goto Cleanup;
	}

	{
		bstr_t bstrClassName(pszClassName);
		hr = pAssembly->GetType_2(bstrClassName, &pClass);
		if (FAILED(hr))
		{
			FormatError(L"Failed to get the type \"%s\" w/hr 0x%08lx\n", pszClassName, hr);
			goto Cleanup;
		}
	}

	return TRUE;
Cleanup:
	Stop();
	return FALSE;
}

void
ClrRuntimeHostV4::Call(PCWSTR method, int argc, intptr_t pArgs[])
{
	ClearError();

	bstr_t bstrMethodName(method);
	SAFEARRAY *psaMethodArgs = NULL;
	variant_t vtEmpty;
	variant_t vtReturn;

    if (argc == 0)
    {
        psaMethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, 0);
    }
    else
    {
        psaMethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, argc);
		/*
        long idx [] = { 0 };
		PutElement(psaMethodArgs, idx, pArg1);
        if (args == 2)
        {
            idx[0] = 1;
			PutElement(psaMethodArgs, idx, pArg2);
        }
        if (args == 3)
        {
            idx[0] = 2;
			PutElement(psaMethodArgs, idx, pArg3);
		}
		*/
    }

	HRESULT hr = pClass->InvokeMember_3(
		bstrMethodName,
		static_cast<BindingFlags>(BindingFlags_InvokeMethod | BindingFlags_Static | BindingFlags_Public),
		NULL,
		vtEmpty, 
		psaMethodArgs,
		&vtReturn);
	if (FAILED(hr))
	{
        FormatError(L"Failed to invoke %s w/hr 0x%08lx\n", method, hr);
		goto Cleanup;
	}
	return;

Cleanup:
	if (psaMethodArgs)
	{
		SafeArrayDestroy(psaMethodArgs);
	}
}

void
ClrRuntimeHostV4::Stop()
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

	if (pClass)
	{
		pClass->Release();
	}

}

void ClrRuntimeHostV4::PutElement(SAFEARRAY* pa, long idx[], PCWSTR pArg)
{
	variant_t varg(pArg);
	VARIANT v = varg;
	SafeArrayPutElement(pa, idx, &v);
}
