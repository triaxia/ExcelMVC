#ifndef Addin_H
#define Addin_H
#include <windows.h>

#define EXPORT_COUNT 10000
typedef void(__stdcall* PFN)();

class AddIn
{
public:
	static HMODULE hModule;
	static PFN Functions[EXPORT_COUNT];
};

#endif //PCH_H
