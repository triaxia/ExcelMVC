#ifndef _ClrRuntimeHostV4_h
#define _ClrRuntimeHostV4_h

#include "ClrRuntimeHost.h"

class ClrRuntimeHostV4 : public ClrRuntimeHost
{
public:
    virtual void Start(PCWSTR pszAssemblyName, PCWSTR pszClassName);
    virtual void Stop();
    virtual void Attach(AddInHead* pHead);
    virtual void Detach();
    virtual void Show();
    virtual void Hide();
    virtual void Click();
    virtual void Run();

private:
    static void Call(int idx, void *arg, bool setArg = false);
};

#endif