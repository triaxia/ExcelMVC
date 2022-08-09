#include "pch.h"
#include "ClrRuntimeHost.h"
#include <cwchar>

WCHAR ClrRuntimeHost::ErrorBuffer[1024] = {};

void
ClrRuntimeHost::FormatError(PCWSTR format, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, hr);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg, hr);
}

BOOL
ClrRuntimeHost::TestAndDisplayError()
{
	BOOL result = wcslen(ErrorBuffer) == 0;
	if (!result) MessageBox(0, ErrorBuffer, L"ExcelMvc", MB_OK + MB_ICONERROR);
	return result;
}

BOOL ClrRuntimeHost::FindAppConfig(PCWSTR basePath, TCHAR *buffer, DWORD size)
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

void ClrRuntimeHost::GetBasePath(TCHAR* buffer, DWORD size)
{
	::GetModuleFileName(Constants::Dll, buffer, size);

	// trim off file name
	int pos = wcslen(buffer);
	while (--pos >= 0 && buffer[pos] != '\\');
	buffer[pos] = 0;
}


