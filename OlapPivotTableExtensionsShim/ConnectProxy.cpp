// ConnectProxy.cpp 
#include "stdafx.h"
#include "CLRLoader.h"
#include "ConnectProxy.h"

// These strings are specific to the managed assembly that this shim will load.
static LPCWSTR szAddInAssemblyName = 
	L"OlapPivotTableExtensions, PublicKeyToken=9805f7e30acd4f18";
static LPCWSTR szConnectClassName = 
	L"OlapPivotTableExtensions.Connect";
static LPCWSTR szAssemblyConfigName =
	L"OlapPivotTableExtensions.dll.config";

CConnectProxy::CConnectProxy() 
    : m_pConnect(NULL), m_pCLRLoader(NULL), m_pUnknownInner(NULL)
{
}

HRESULT CConnectProxy::FinalConstruct()
{
    HRESULT hr = S_OK;
    IUnknown* pUnkThis = NULL;

    // Instantiate the CLR-loader object.
    m_pCLRLoader = new (std::nothrow) CCLRLoader();
    IfNullGo( m_pCLRLoader );

    IfFailGo( this->QueryInterface(IID_IUnknown, (LPVOID*)&pUnkThis) );

    // Load the CLR, create an AppDomain, and instantiate the target add-in
    // and the inner aggregated object of the shim.
    IfFailGo( m_pCLRLoader->CreateAggregatedAddIn(
        pUnkThis,
        szAddInAssemblyName, szConnectClassName, szAssemblyConfigName) );

    // Extract the IDTExtensibility2 interface pointer from the target add-in.
    IfFailGo( m_pUnknownInner->QueryInterface(
        __uuidof(IDTExtensibility2), (LPVOID*)&this->m_pConnect) );

Error:
    if (pUnkThis != NULL)
        pUnkThis->Release();

    return hr;
}

// Cache the pointer to the aggregated innner object, and make sure
// we increment the refcount on it.
HRESULT __stdcall CConnectProxy::SetInnerPointer(IUnknown* pUnkInner)
{
    if (pUnkInner == NULL)
    {
        return E_POINTER;
    }
    if (m_pUnknownInner != NULL)
    {
        return E_UNEXPECTED;
    }
    
    m_pUnknownInner = pUnkInner;
    m_pUnknownInner->AddRef();
    return S_OK;
}

// IDTExtensibility2 implementation: OnConnection, OnAddInsUpdate and
// OnStartupComplete are simple pass-throughs to the proxied managed
// add-in. We only need to wrap IDTExtensibility2 because we need to
// add behavior to the OnBeginShutdown and OnDisconnection methods.
HRESULT __stdcall CConnectProxy::OnConnection(
    IDispatch * Application, ext_ConnectMode ConnectMode, 
    IDispatch *AddInInst, SAFEARRAY **custom)
{
    return m_pConnect->OnConnection(
        Application, ConnectMode, AddInInst, custom);
}

HRESULT __stdcall CConnectProxy::OnAddInsUpdate(SAFEARRAY **custom)
{
    return m_pConnect->OnAddInsUpdate(custom);
}

HRESULT __stdcall CConnectProxy::OnStartupComplete(SAFEARRAY **custom)
{
    return m_pConnect->OnStartupComplete(custom);
}

// When the host application shuts down, it calls OnBeginShutdown, 
// and then OnDisconnection. We must be careful to test that the add-in
// pointer is not null, to allow for the case where the add-in was
// previously disconnected prior to app shutdown.
HRESULT __stdcall CConnectProxy::OnBeginShutdown(SAFEARRAY **custom)
{
    HRESULT hr = S_OK;
    if (m_pConnect)
    {
        hr = m_pConnect->OnBeginShutdown(custom);
    }
    return hr;
}

// OnDisconnection is called if the user disconnects the add-in via the COM
// add-ins dialog. We wrap this so that we can make sure we can clean up
// the reference we're holding to the inner object. We must also allow for 
// the possibility that the user has disconnected the add-in via the COM 
// add-ins dialog or programmatically: in this scenario, OnDisconnection is
// called first, and this add-in never gets the OnBeginShutdown call
// (because it has already been disconnected by then).
HRESULT __stdcall CConnectProxy::OnDisconnection(
    ext_DisconnectMode RemoveMode, SAFEARRAY **custom)
{
	HRESULT hr = S_OK;
    hr =  m_pConnect->OnDisconnection(RemoveMode, custom);
    if (SUCCEEDED(hr))
    {
        m_pConnect->Release();
        m_pConnect = NULL;
    }


	// FIX: Bug discovered after release of 2.3.1.0.
	// Move the code that unloads the AppDomain 
	// from the CConnectProxy::FinalRelease to 
	// CConnectProxy::OnDisconnection.
	if (m_pCLRLoader)
	{
		m_pCLRLoader->Unload();
		delete m_pCLRLoader;
		m_pCLRLoader = NULL;
	}

	return hr;
}

// Make sure we unload the AppDomain, and clean up our references. 
// FinalRelease will be the last thing called in the shim/add-in, after
// OnBeginShutdown and OnDisconnection.
void CConnectProxy::FinalRelease() 
{
    // Release the aggregated inner object.
    if (m_pUnknownInner)
    {
        m_pUnknownInner->Release();
    }
}
