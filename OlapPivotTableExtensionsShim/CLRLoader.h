// CLRLoader.h
#pragma once

class CCLRLoader
{
public:
    CCLRLoader(void);
    virtual ~CCLRLoader(void);

    HRESULT CreateAggregatedAddIn(
        IUnknown* pOuter,
        LPCWSTR szAssemblyName, 
        LPCWSTR szClassName, 
        LPCWSTR szAssemblyConfigName);
    HRESULT Unload(void);

private:
    HRESULT LoadCLR();
    HRESULT CreateAppDomain(LPCWSTR szAssemblyConfigName);
    HRESULT GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize);

    ICorRuntimeHost *m_pCorRuntimeHost;
    mscorlib::_AppDomain *m_pAppDomain;
};
