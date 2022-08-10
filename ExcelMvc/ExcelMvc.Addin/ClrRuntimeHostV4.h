#ifndef _ClrRuntimeHostV4_h
#define _ClrRuntimeHostV4_h

#include "ClrRuntimeHost.h"

class ClrRuntimeHostV4 : public ClrRuntimeHost
{
public:
    virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName);
    virtual void Stop();
    virtual void CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName,
        PCWSTR pArg1 = NULL, PCWSTR pArg2 = NULL, PCWSTR pArg3 = NULL);
private:
    static void PutElement(SAFEARRAY* pa, long idx[], PCWSTR pArg);
};

#endif