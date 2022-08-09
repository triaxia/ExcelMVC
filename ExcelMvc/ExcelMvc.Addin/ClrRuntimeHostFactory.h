#ifndef _ClrRuntimeHostFactory_h
#define _ClrRuntimeHostFactory_h
#include "ClrRuntimeHost.h"

class ClrRuntimeHostFactory
{
public:
	static ClrRuntimeHost* Create();
};

#endif