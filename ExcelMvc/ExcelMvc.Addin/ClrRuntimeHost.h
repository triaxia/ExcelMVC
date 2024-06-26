#ifndef _ClrRuntimeHost_h
#define _ClrRuntimeHost_h

#include <string>
using string_t = std::basic_string<wchar_t>;

struct AddInHead;

class ClrRuntimeHost
{
public:
    virtual void Start(PCWSTR pszAssemblyName, PCWSTR pszClassName) = 0;
    virtual void Stop() = 0;
    virtual void Attach(AddInHead*pHead) = 0;
    virtual void Detach() = 0;
    virtual void Show() = 0;
    virtual void Hide() = 0;
    virtual void Click() = 0;
    virtual void Run() = 0;

    static BOOL TestAndDisplayError();
    static BOOL FindFile(PCWSTR basePath, PCWSTR pattern, TCHAR* buffer, DWORD size);
    static string_t GetModulePath();
    static string_t GetModuleName();
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