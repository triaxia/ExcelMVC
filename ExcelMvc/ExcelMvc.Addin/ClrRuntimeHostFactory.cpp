#include "pch.h"
#include "ClrRuntimeHostFactory.h"
#include "ClrRuntimeHostV4.h"
#include "ClrRuntimeHostCore.h"
#include <filesystem>

ClrRuntimeHost* ClrRuntimeHostFactory ::Create()
{
	std::filesystem::path p(ClrRuntimeHost::GetRuntimeConfigFile().c_str());
	if (std::filesystem::exists(p))
		return new ClrRuntimeHostCore();
	else
		return new ClrRuntimeHostV4();
}


