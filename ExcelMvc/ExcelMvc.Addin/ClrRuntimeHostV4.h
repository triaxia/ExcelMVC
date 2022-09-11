#ifndef _ClrRuntimeHostV4_h
#define _ClrRuntimeHostV4_h

#include "ClrRuntimeHost.h"

class ClrRuntimeHostV4 : public ClrRuntimeHost
{
public:
    virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName,
        PCWSTR pszClassName, int argc, PCWSTR methods[]);
    virtual void Stop();
    virtual void Call(PCWSTR method, int argc, intptr_t pArgs[]);
private:
    static void PutElement(SAFEARRAY* pa, long idx[], PCWSTR pArg);
};

#endif