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

static ICLRMetaHost* pMetaHost = NULL;
static ICLRRuntimeInfo* pRuntimeInfo = NULL;

// ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting interfaces
// supported by CLR 4.0. Here we demo the ICorRuntimeHost interface that 
// was provided in .NET v1.x, and is compatible with all .NET Frameworks. 
static ICorRuntimeHost* pCorRuntimeHost = NULL;

// ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting interfaces
// supported by CLR 4.0. Here we demo the ICLRRuntimeHost interface that 
// was provided in .NET v2.0 to support CLR 2.0 new features. 
// ICLRRuntimeHost does not support loading the .NET v1.x runtimes.
static ICLRRuntimeHost* pClrRuntimeHost = NULL;

static IUnknownPtr pAppDomainSetupThunk = NULL;
static IAppDomainSetupPtr pAppDomainSetup = NULL;

static IUnknownPtr pAppDomainThunk = NULL;
static _AppDomainPtr pAppDomain = NULL;

static _AssemblyPtr pAssembly = NULL;

static _TypePtr pClass = NULL;

static LPCTSTR pVersion = L"v4.0.30319";

static LPCWSTR MethodNames[] =
{
	L"Attach",
	L"Detach",
	L"Show",
	L"Hide",
	L"Click",
	L"Run"
};

void
ClrRuntimeHostV4::Start(PCWSTR pszAssemblyName, PCWSTR pszClassName)
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
	hr = pMetaHost->GetRuntime(pVersion, IID_PPV_ARGS(&pRuntimeInfo));
	if (FAILED(hr))
	{
		FormatError(L"ICLRMetaHost::GetRuntime (%s) failed w/hr 0x%08lx\n", pVersion, hr);
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
		FormatError(L".NET runtime %s cannot be loaded\n", pVersion);
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

	// set app config file if there is one matching *.xll.config or *.dll.config 
	// in the base path
	TCHAR configFile[MAX_PATH];
	if (FindAppConfig(basePath.c_str(), L"*.xll.config", configFile, MAX_PATH)
	 || FindAppConfig(basePath.c_str(), L"*.dll.config", configFile, MAX_PATH))
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

	/*
	{
		auto argc = sizeof(MethodNames) / sizeof(LPCWSTR);
		for (auto idx = 0; idx < argc; idx++)
		{
			bstr_t bstrMethodName(MethodNames[idx]);
			SAFEARRAY* psaResult = SafeArrayCreateVector(VT_VARIANT, 0, 1);
			hr = pClass->GetMember(bstrMethodName, MemberTypes_Method, static_cast<BindingFlags>(BindingFlags_Static | BindingFlags_Public), &psaResult);
			if (FAILED(hr))
			{
				FormatError(L"Failed to get the method \"%s\" w/hr 0x%08lx\n", MethodNames[idx], hr);
				goto Cleanup;
			}
			long lowerBound, upperBound;  // get array bounds
			SafeArrayGetLBound(psaResult, 0, &lowerBound);
			SafeArrayGetUBound(psaResult, 0, &upperBound);

			SafeArrayGetLBound(psaResult, 1, &lowerBound);
			SafeArrayGetUBound(psaResult, 1, &upperBound);
			long dim[] = { 0 };
			VARIANT v;
			SafeArrayGetElement(psaResult, dim, &v);
			//Functions[idx] = function;
			SafeArrayDestroy(psaResult);
		};
	}
	*/

	return;
Cleanup:
	Stop();
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

void
ClrRuntimeHostV4::Attach(AddInHead* pHead)
{
	Call(0, pHead, true);
}

void
ClrRuntimeHostV4::Detach()
{
	Call(1, NULL);
}

void
ClrRuntimeHostV4::Show()
{
	Call(2, NULL);
}

void
ClrRuntimeHostV4::Hide()
{
	Call(3, NULL);
}

void
ClrRuntimeHostV4::Click()
{
	Call(4, NULL);
}

void
ClrRuntimeHostV4::Run()
{
	Call(5, NULL);
}

void
ClrRuntimeHostV4::Call(int idx, void* arg, bool setArg)
{
	ClearError();

	bstr_t bstrMethodName(MethodNames[idx]);
	SAFEARRAY* psaMethodArgs = NULL;
	variant_t vtEmpty;
	variant_t vtReturn;
	
	psaMethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, setArg ? 1 : 0);
	if (setArg)
	{
		variant_t vtArg((intptr_t)arg);
		long index = 0;
		SafeArrayPutElement(psaMethodArgs, &index, &vtArg);
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
		FormatError(L"Failed to invoke %s w/hr 0x%08lx\n", MethodNames[idx], hr);
		goto Cleanup;
	}
	return;

Cleanup:
	if (psaMethodArgs)
	{
		SafeArrayDestroy(psaMethodArgs);
	}
}
