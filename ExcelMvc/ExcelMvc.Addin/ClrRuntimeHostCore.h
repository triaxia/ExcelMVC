#ifndef _ClrRuntimeHostCore_h
#define _ClrRuntimeHostCore_h

#include "ClrRuntimeHost.h"
#include <XLCALL.H>

class ClrRuntimeHostCore : public ClrRuntimeHost
{
public:
	virtual BOOL Start(PCWSTR pszAssemblyName, PCWSTR pszClassName
		, int argc, PCWSTR methods[]);
	virtual void Stop();
	virtual void Call(int idx, int argc, void* args[]);

private:
	std::map<int, void *> Functions;
};

#endif