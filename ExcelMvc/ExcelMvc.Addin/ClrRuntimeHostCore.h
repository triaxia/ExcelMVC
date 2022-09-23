#ifndef _ClrRuntimeHostCore_h
#define _ClrRuntimeHostCore_h

#include "ClrRuntimeHost.h"

class ClrRuntimeHostCore : public ClrRuntimeHost
{
public:
	virtual BOOL Start(PCWSTR pszAssemblyName, PCWSTR pszClassName
		, int argc, PCWSTR methods[]);
	virtual void Stop();
	virtual void Call(PCWSTR method, int argc, intptr_t pArgs[]);

private:
	std::map<PCWSTR, void *> Functions;
};

#endif