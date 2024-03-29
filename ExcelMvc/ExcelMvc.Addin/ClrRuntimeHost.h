#ifndef _ClrRuntimeHost_h
#define _ClrRuntimeHost_h

#include <string>
using string_t = std::basic_string<wchar_t>;

class ClrRuntimeHost
{
public:
    virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName) = 0;
    virtual void Stop() = 0;
    virtual void CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName,
        PCWSTR pArg1 = NULL, PCWSTR pArg2 = NULL, PCWSTR pArg3 = NULL) = 0;

    static BOOL TestAndDisplayError();
    static BOOL FindAppConfig(PCWSTR basePath, TCHAR* buffer, DWORD size);
    static string_t GetBasePath();
    static string_t GetRuntimeConfigFile();

protected:
    static void FormatError(PCWSTR format, HRESULT hr);
    static void FormatError(PCWSTR format, PCWSTR arg, HRESULT hr);
    static void FormatError(PCWSTR format, PCWSTR arg1, PCWSTR arg2, HRESULT hr);
    static void FormatError(PCWSTR format, PCWSTR arg);
    static void FormatError(PCWSTR format, PCWSTR arg1, PCWSTR arg2);
    static void ClearError();
private:
    static WCHAR ErrorBuffer[1024];
};

#endif