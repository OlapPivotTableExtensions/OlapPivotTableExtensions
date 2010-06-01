#include "StdAfx.h"
#include "clrloader.h"

using namespace mscorlib;

CCLRLoader::CCLRLoader(void) 
    : m_pCorRuntimeHost(NULL), m_pAppDomain(NULL)
{
}

// CreateInstance: loads the CLR, creates an AppDomain, and creates an 
// aggregated instance of the target managed add-in in that AppDomain.
HRESULT CCLRLoader::CreateAggregatedAddIn(
    IUnknown* pOuter,
    LPCWSTR szAssemblyName, 
    LPCWSTR szClassName, 
    LPCWSTR szAssemblyConfigName)
{
    HRESULT hr = E_FAIL;

    CComPtr<_ObjectHandle>                              srpObjectHandle;
    CComPtr<ManagedHelpers::IManagedAggregator >        srpManagedAggregator;
    CComPtr<IComAggregator>                             srpComAggregator;
    CComVariant                                         cvarManagedAggregator;
 
    // Load the CLR, and create an AppDomain for the target assembly.
    IfFailGo( LoadCLR() );
    IfFailGo( CreateAppDomain(szAssemblyConfigName) );

    // Create the managed aggregator in the target AppDomain, and unwrap it.
	// This component needs to be in a location where fusion will find it, ie
	// either in the GAC or in the same folder as the shim and the add-in.
    IfFailGo( m_pAppDomain->CreateInstance(
        CComBSTR(L"ManagedAggregator, PublicKeyToken=9805f7e30acd4f18"),
        CComBSTR(L"ManagedHelpers.ManagedAggregator"),
        &srpObjectHandle) );
    IfFailGo( srpObjectHandle->Unwrap(&cvarManagedAggregator) );
    IfFailGo( cvarManagedAggregator.pdispVal->QueryInterface(
        &srpManagedAggregator) );

    // Instantiate and aggregate the inner managed add-in into the outer
    // (unmanaged, ConnectProxy) object.
    IfFailGo( pOuter->QueryInterface(
        __uuidof(IComAggregator), (LPVOID*)&srpComAggregator) );
    IfFailGo( srpManagedAggregator->CreateAggregatedInstance(
        CComBSTR(szAssemblyName), CComBSTR(szClassName), srpComAggregator) );

Error:
    return hr;
}

// LoadCLR: loads and starts the .NET CLR.
HRESULT CCLRLoader::LoadCLR()
{
    HRESULT hr = S_OK;

    // Ensure the CLR is only loaded once.
    if (m_pCorRuntimeHost != NULL)
    {
        return hr;
    }

    // Load the CLR into the process, using the default (latest) version, 
    // the default ("wks") flavor, and default (single) domain.
    hr = CorBindToRuntimeEx(
        0, 0, 0, 
        CLSID_CorRuntimeHost, IID_ICorRuntimeHost, 
		(LPVOID*)&m_pCorRuntimeHost);

    // If CorBindToRuntimeEx returned a failure HRESULT, we failed to 
	// load the CLR.
    if (!SUCCEEDED(hr)) 
    {
        return hr;
    }

    // Start the CLR.
    return m_pCorRuntimeHost->Start();
}

// In order to securely load an assembly, its fully qualified strong name
// and not the filename must be used. To do that, the target AppDomain's 
// base directory needs to point to the directory where the assembly is.
HRESULT CCLRLoader::CreateAppDomain(LPCWSTR szAssemblyConfigName)
{
    USES_CONVERSION;
    HRESULT hr = S_OK;

    // Ensure the AppDomain is created only once.
    if (m_pAppDomain != NULL)
    {
        return hr;
    }

    CComPtr<IUnknown> pUnkDomainSetup;
    CComPtr<IAppDomainSetup> pDomainSetup;
    CComPtr<IUnknown> pUnkAppDomain;
    TCHAR szDirectory[MAX_PATH + 1];
    TCHAR szAssemblyConfigPath[MAX_PATH + 1];
    CComBSTR cbstrAssemblyConfigPath;

    // Create an AppDomainSetup with the base directory pointing to the
    // location of the managed DLL. We assume that the target assembly
    // is located in the same directory.
    IfFailGo( m_pCorRuntimeHost->CreateDomainSetup(&pUnkDomainSetup) );
    IfFailGo( pUnkDomainSetup->QueryInterface(
        __uuidof(pDomainSetup), (LPVOID*)&pDomainSetup) );

    // Get the location of the hosting shim DLL, and configure the 
    // AppDomain to search for assemblies in this location.
    IfFailGo( GetDllDirectory(
        szDirectory, sizeof(szDirectory)/sizeof(szDirectory[0])) );
    pDomainSetup->put_ApplicationBase(CComBSTR(szDirectory));

    // Set the AppDomain to use a local DLL config if there is one.
    IfFailGo( StringCchCopy(
        szAssemblyConfigPath, 
        sizeof(szAssemblyConfigPath)/sizeof(szAssemblyConfigPath[0]), 
        szDirectory) );
    if (!PathAppend(szAssemblyConfigPath, szAssemblyConfigName))
    {
        hr = E_UNEXPECTED;
        goto Error;
    }
    IfFailGo( cbstrAssemblyConfigPath.Append(szAssemblyConfigPath) );
    IfFailGo( pDomainSetup->put_ConfigurationFile(cbstrAssemblyConfigPath) );

    // Create an AppDomain that will run the managed assembly, and get the
    // AppDomain's _AppDomain pointer from its IUnknown pointer.
    IfFailGo( m_pCorRuntimeHost->CreateDomainEx(T2W(szDirectory), 
        pUnkDomainSetup, 0, &pUnkAppDomain) );
    IfFailGo( pUnkAppDomain->QueryInterface(
        __uuidof(m_pAppDomain), (LPVOID*)&m_pAppDomain) );

Error:
   return hr;
}

// GetDllDirectory: gets the directory location of the DLL containing this
// code - that is, the shim DLL. The target add-in DLL will also be in this
// directory.
HRESULT CCLRLoader::GetDllDirectory(TCHAR *szPath, DWORD nPathBufferSize)
{
    // Get the shim DLL module instance, or bail.
    HMODULE hInstance = _AtlBaseModule.GetModuleInstance();
    if (hInstance == 0)
    {
        return E_FAIL;
    }

    // Get the shim DLL filename, or bail.
    TCHAR szModule[MAX_PATH + 1];
    DWORD dwFLen = ::GetModuleFileName(hInstance, szModule, MAX_PATH);
    if (dwFLen == 0)
    {
        return E_FAIL;
    }

    // Get the full path to the shim DLL, or bail.
    TCHAR *pszFileName;
    dwFLen = ::GetFullPathName(
        szModule, nPathBufferSize, szPath, &pszFileName);
    if (dwFLen == 0 || dwFLen >= nPathBufferSize)
    {
        return E_FAIL;
    }

    *pszFileName = 0;
    return S_OK;
}

// Unload the AppDomain. This will be called by the ConnectProxy
// in OnDisconnection.
HRESULT CCLRLoader::Unload(void)
{
    HRESULT hr = S_OK;
    IUnknown* pUnkDomain = NULL;
    IfFailGo(m_pAppDomain->QueryInterface(
        __uuidof(IUnknown), (LPVOID*)&pUnkDomain));
    hr = m_pCorRuntimeHost->UnloadDomain(pUnkDomain);

    // Added in 2.0.2.0, only for Add-ins.
    m_pAppDomain->Release();
    m_pAppDomain = NULL;
    
Error:
    if (pUnkDomain != NULL)
    {
        pUnkDomain->Release();
    }
    return hr;
}

CCLRLoader::~CCLRLoader(void)
{
    if (m_pAppDomain)
    {
        m_pAppDomain->Release();
    }
    if (m_pCorRuntimeHost)
    {
        m_pCorRuntimeHost->Release();
    }
}
