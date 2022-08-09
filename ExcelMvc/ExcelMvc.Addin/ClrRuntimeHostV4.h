#ifndef _ClrRuntimeHostV4_h
#define _ClrRuntimeHostV4_h

#include "ClrRuntimeHost.h"

class ClrRuntimeHostV4 : public ClrRuntimeHost
{
public:
    virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName);
    virtual void Stop();
    virtual void CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName,
        VARIANT *pArg1 = NULL, VARIANT *pArg2 = NULL, VARIANT *pArg3 = NULL);
};

#endif