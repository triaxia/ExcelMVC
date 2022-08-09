
#include "pch.h"
#include "ClrRuntimeHostCore.h"

#define NETHOST_USE_AS_STATIC
#include <nethost.h>

#include <coreclr_delegates.h>
#include <hostfxr.h>
#include <string>

using string_t = std::basic_string<char_t>;

hostfxr_initialize_for_runtime_config_fn init_fptr;
hostfxr_get_runtime_delegate_fn get_delegate_fptr;
hostfxr_close_fn close_fptr;
load_assembly_and_get_function_pointer_fn load_fptr;

void* load_library(const char_t* path)
{
	HMODULE h = ::LoadLibraryW(path);
	return (void*)h;
}

void* get_export(void* h, const char* name)
{
	void* f = ::GetProcAddress((HMODULE)h, name);
	return f;
}

bool load_hostfxr()
{
	// Pre-allocate a large buffer for the path to hostfxr
	char_t buffer[MAX_PATH];
	size_t buffer_size = sizeof(buffer) / sizeof(char_t);
	int rc = get_hostfxr_path(buffer, &buffer_size, nullptr);
	if (rc != 0) return false;

	// Load hostfxr and get desired exports
	void* lib = load_library(buffer);
	init_fptr = (hostfxr_initialize_for_runtime_config_fn)get_export(lib, "hostfxr_initialize_for_runtime_config");
	get_delegate_fptr = (hostfxr_get_runtime_delegate_fn)get_export(lib, "hostfxr_get_runtime_delegate");
	close_fptr = (hostfxr_close_fn)get_export(lib, "hostfxr_close");

	return (init_fptr && get_delegate_fptr && close_fptr);
}

load_assembly_and_get_function_pointer_fn get_dotnet_load_assembly(const char_t* config_path)
{
	// Load .NET Core
	void* load_assembly_and_get_function_pointer = nullptr;
	hostfxr_handle cxt = nullptr;
	int rc = init_fptr(config_path, nullptr, &cxt);
	if (rc != 0 || cxt == nullptr)
	{
		close_fptr(cxt);
		return nullptr;
	}

	// Get the load assembly function pointer
	rc = get_delegate_fptr(
		cxt,
		hdt_load_assembly_and_get_function_pointer,
		&load_assembly_and_get_function_pointer);

	close_fptr(cxt);
	return (load_assembly_and_get_function_pointer_fn)load_assembly_and_get_function_pointer;
}

string_t AssemblyName;
string_t BasePath;

BOOL
ClrRuntimeHostCore::Start(PCWSTR pszVersion, PCWSTR pszAssemblyName)
{
	AssemblyName = pszAssemblyName;
	WCHAR buffer[MAX_PATH];
	GetBasePath(buffer, sizeof(buffer) / sizeof(WCHAR));
	BasePath = buffer;

	if (!load_hostfxr())
	{
		FormatError(L"%s failed", L"load_hostfxr");
		return FALSE;
	}

	const string_t config_path = BasePath + L"\\test.runtimeconfig.json";
	load_fptr = get_dotnet_load_assembly(config_path.c_str());
	if (load_fptr == nullptr)
	{
		FormatError(L"%s failed (%s)", config_path.c_str());
		return FALSE;
	}

	return TRUE;
}

void
ClrRuntimeHostCore::CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName, VARIANT* pArg1, VARIANT* pArg2, VARIANT* pArg3)
{
	const string_t dotnetlib_path = BasePath + +L"\\" + AssemblyName + L".dll";
	const string_t dotnet_type = string_t(pszClassName) + L",ExcelMvc";
	component_entry_point_fn method = nullptr;
	int rc = load_fptr(
		dotnetlib_path.c_str(),
		dotnet_type.c_str(),
		pszMethodName,
		nullptr /*delegate_type_name*/,
		nullptr,
		(void**)&method);
	if (method == nullptr)
	{
		FormatError(L"%s not found (%s)", dotnet_type.c_str());
		return;
	}

	struct lib_args {};
	lib_args args;
	method(&args, sizeof(args));
}

void
ClrRuntimeHostCore::Stop()
{
}