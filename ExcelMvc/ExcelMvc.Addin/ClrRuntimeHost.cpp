#include "pch.h"
#include "ClrRuntimeHost.h"
#include <cwchar>

extern "C" { extern WCHAR ModuleFileName[]; }

WCHAR ClrRuntimeHost::ErrorBuffer[1024] = {};

void
ClrRuntimeHost::ClearError()
{
    ErrorBuffer[0] = 0;
}

void
ClrRuntimeHost::FormatError(PCWSTR format, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, hr);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg, hr);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg1, PCWSTR arg2, HRESULT hr)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg1, arg2, hr);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg);
}

void
ClrRuntimeHost::FormatError(PCWSTR format, PCWSTR arg1, PCWSTR arg2)
{
    swprintf(ErrorBuffer, sizeof(ErrorBuffer) / sizeof(WCHAR), format, arg1, arg2);
}

BOOL
ClrRuntimeHost::TestAndDisplayError()
{
	BOOL result = wcslen(ErrorBuffer) == 0;
	if (!result) MessageBox(0, ErrorBuffer, L"ExcelMvc", MB_OK + MB_ICONERROR);
	return result;
}

BOOL ClrRuntimeHost::FindFile(PCWSTR basePath, PCWSTR pattern, TCHAR *buffer, DWORD size)
{
    TCHAR file[MAX_PATH];
    swprintf(file, MAX_PATH, L"%s\\%s", basePath, pattern);

    WIN32_FIND_DATA data;
    HANDLE hfile = ::FindFirstFile(file, &data);
    if (hfile != INVALID_HANDLE_VALUE)
    {
        swprintf(buffer, size, L"%s\\%s", basePath, data.cFileName);
        FindClose(hfile);
        return true;
    }
    return false;
}

string_t ClrRuntimeHost::GetModulePath()
{
    string_t path = ModuleFileName;
    auto pos = path.find_last_of(L"\\");
    return path.substr(0, pos);
}

string_t ClrRuntimeHost::GetModuleName()
{
    string_t path = ModuleFileName;
    auto pos = path.find_last_of(L"\\");
    return path.substr(pos + 1);
}

string_t ClrRuntimeHost::GetRuntimeConfigFile()
{
    auto name = GetModuleName();
    
    // remove extension
    auto pos = name.find_last_of(L".");
    if (pos != string_t::npos)
        name = name.substr(0, pos);

    // remove trailing 64
    pos = name.find_last_of(L"64");
    if (pos != string_t::npos && pos == name.size() - 1)
        name = name.substr(0, pos - 1);

    return GetModulePath() + L"\\" + name + L".runtimeconfig.json";
}