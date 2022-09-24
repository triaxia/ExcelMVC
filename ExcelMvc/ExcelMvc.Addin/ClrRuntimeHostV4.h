#ifndef _ClrRuntimeHostV4_h
#define _ClrRuntimeHostV4_h

#include "ClrRuntimeHost.h"

class ClrRuntimeHostV4 : public ClrRuntimeHost
{
public:
    virtual void Start(PCWSTR pszAssemblyName, PCWSTR pszClassName);
    virtual void Stop();
    virtual void Attach();
    virtual void Detach();
    virtual void Show();
    virtual void Hide();
    virtual void Click();
    virtual void Run();
    virtual void Udf(int argc, void* args[]);

private:
    static void PutElement(SAFEARRAY* pa, long idx[], PCWSTR pArg);
};

#endif