#ifndef _ClrRuntimeHostCore_h
#define _ClrRuntimeHostCore_h

#include "ClrRuntimeHost.h"
#include <XLCALL.H>

class ClrRuntimeHostCore : public ClrRuntimeHost
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

private:
	std::map<int, void *> Functions;
};

#endif