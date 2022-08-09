class ClrRuntimeHostCore : public ClrRuntimeHost
{
public:
	virtual BOOL Start(PCWSTR pszVersion, PCWSTR pszAssemblyName);
	virtual void Stop();
	virtual void CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName,
        VARIANT *pArg1 = NULL, VARIANT *pArg2 = NULL, VARIANT *pArg3 = NULL);
};