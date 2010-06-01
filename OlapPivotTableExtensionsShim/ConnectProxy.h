// ConnectProxy.h
#pragma once
#include "resource.h"
#include "OlapPivotTableExtensionsShim.h"

using namespace AddInDesignerObjects;

class ATL_NO_VTABLE CConnectProxy :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CConnectProxy, &CLSID_ConnectProxy>,
    public IDispatchImpl<_IDTExtensibility2,
        &IID__IDTExtensibility2,
        &LIBID_AddInDesignerObjects, 1, 0>,
    public IComAggregator
{
public:
    CConnectProxy();

    DECLARE_REGISTRY_RESOURCEID(IDR_CONNECTPROXY)
    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CConnectProxy)
        COM_INTERFACE_ENTRY(IDTExtensibility2)
        COM_INTERFACE_ENTRY(IComAggregator)
        COM_INTERFACE_ENTRY_AGGREGATE_BLIND(m_pUnknownInner)
    END_COM_MAP()

    HRESULT FinalConstruct();
    void FinalRelease();

public:
    //IDTExtensibility2.
    STDMETHOD(OnConnection)(
        IDispatch * Application, ext_ConnectMode ConnectMode,
        IDispatch *AddInInst, SAFEARRAY **custom);

    STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom);

    STDMETHOD(OnStartupComplete)(SAFEARRAY **custom);

    STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom);

    STDMETHOD(OnDisconnection)(
        ext_DisconnectMode RemoveMode, SAFEARRAY **custom);

    // IComAggregator.
    STDMETHOD(SetInnerPointer)(IUnknown* pUnkInner);

private:
    IDTExtensibility2 *m_pConnect;
    CCLRLoader *m_pCLRLoader;
    IUnknown *m_pUnknownInner;
};

OBJECT_ENTRY_AUTO(__uuidof(ConnectProxy), CConnectProxy)
