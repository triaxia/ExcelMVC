#ifndef _ClrRuntimeHost_h
#define _ClrRuntimeHost_h

class ClrRuntimeHost
{
public:
    virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName) = 0;
    virtual void Stop() = 0;
    virtual void CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName,
        VARIANT *pArg1 = NULL, VARIANT *pArg2 = NULL, VARIANT *pArg3 = NULL) = 0;
    static BOOL TestAndDisplayError();

protected:
    static WCHAR ErrorBuffer[1024];

    static void FormatError(PCWSTR format, HRESULT hr);
    static void FormatError(PCWSTR format, PCWSTR arg, HRESULT hr);
    static void FormatError(PCWSTR format, PCWSTR arg1, PCWSTR arg2, HRESULT hr);

    static void FormatError(PCWSTR format, PCWSTR arg);
    static void FormatError(PCWSTR format, PCWSTR arg1, PCWSTR arg2);

    static BOOL FindAppConfig(PCWSTR basePath, TCHAR *buffer, DWORD size);
    static void GetBasePath(TCHAR* buffer, DWORD size);
};

#endif