#include "pch.h"
#include "ClrRuntimeHostFactory.h"
#include "ClrRuntimeHostV4.h"
#include "ClrRuntimeHostCore.h"

ClrRuntimeHost* ClrRuntimeHostFactory ::Create()
{
	return new ClrRuntimeHostCore();
}


