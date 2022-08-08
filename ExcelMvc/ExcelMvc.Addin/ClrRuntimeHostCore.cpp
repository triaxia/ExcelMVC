/****************************** Module Header ******************************\
* Module Name:  ClrRuntimeHost.Core.cpp
* Copyright (c) Microsoft Corporation.
*
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/en-us/openness/licenses.aspx#MPL.
* All other rights reserved.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "pch.h"
#include <windows.h>
#include <metahost.h>
#include "ClrRuntimeHostCore.h"


#pragma region Includes and Imports
#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only				\
	high_property_prefixes("_get", "_put", "_putref")		\
	rename("ReportEvent", "InteropServices_ReportEvent")
using namespace mscorlib;
#pragma endregion

WCHAR ClrRuntimeHostCore::ErrorBuffer[1024] = {};

void
ClrRuntimeHostCore::FormatError(PCWSTR format, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, hr);
}

void
ClrRuntimeHostCore::FormatError(PCWSTR format, PCWSTR arg)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg);
}

void
ClrRuntimeHostCore::FormatError(PCWSTR format, PCWSTR arg, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg, hr);
}

BOOL 
ClrRuntimeHostCore::Start(PCWSTR pszVersion, PCWSTR pszAssemblyName, PCWSTR basePath)
{
	return FALSE;
}

void
ClrRuntimeHostCore::CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName, VARIANT *pArg1, VARIANT *pArg2, VARIANT *pArg3)
{
}

void
ClrRuntimeHostCore::Stop()
{
}

BOOL
ClrRuntimeHostCore::TestAndDisplayError()
{
	BOOL result = wcslen(ErrorBuffer) == 0;
	if (!result)
        MessageBox(0, ErrorBuffer, L"ExcelMvc", MB_OK + MB_ICONERROR);
	return result;
}

BOOL ClrRuntimeHostCore::FindAppConfig(PCWSTR basePath, TCHAR *buffer, DWORD size)
{
    TCHAR pattern[MAX_PATH];
    swprintf(pattern, MAX_PATH, L"%s\\*.dll.config", basePath);

    WIN32_FIND_DATA data;
    HANDLE hfile = ::FindFirstFile(pattern, &data);
    if (hfile != NULL)
    {
        swprintf(buffer, size, L"%s\\%s", basePath, data.cFileName);
        FindClose(hfile);
        return true;
    }
    return false;
}



