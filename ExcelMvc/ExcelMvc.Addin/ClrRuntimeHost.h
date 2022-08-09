class ClrRuntimeHost
{
public:
    virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName);
    virtual void Stop();
    virtual void CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName,
        VARIANT *pArg1 = NULL, VARIANT *pArg2 = NULL, VARIANT *pArg3 = NULL);
    static BOOL TestAndDisplayError();
protected:
    static WCHAR ErrorBuffer[1024];
    static void FormatError(PCWSTR format, HRESULT hr);
    static void FormatError(PCWSTR format, PCWSTR arg, HRESULT hr);
    static void FormatError(PCWSTR format, PCWSTR arg);
    static BOOL FindAppConfig(PCWSTR basePath, TCHAR *buffer, DWORD size);
    static void GetBasePath(TCHAR* buffer, DWORD size);
};