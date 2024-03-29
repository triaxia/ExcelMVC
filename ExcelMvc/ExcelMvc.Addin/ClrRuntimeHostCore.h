#ifndef _ClrRuntimeHostCore_h
#define _ClrRuntimeHostCore_h

#include "ClrRuntimeHost.h"

class ClrRuntimeHostCore : public ClrRuntimeHost
{
public:
	virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName);
	virtual void Stop();
	virtual void CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName,
		PCWSTR pArg1 = NULL, PCWSTR pArg2 = NULL, PCWSTR pArg3 = NULL);
};

#endif