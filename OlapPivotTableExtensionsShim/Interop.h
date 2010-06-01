#pragma once

__interface __declspec(uuid("7b70c487-b741-4973-b915-b812a91bdf63"))
IComAggregator : public IUnknown
{
      HRESULT __stdcall SetInnerPointer(IUnknown *pUnkInner );
};

namespace ManagedHelpers
{
    __interface __declspec(uuid("142a261b-1550-4849-b109-715aa4629a14"))
    IManagedAggregator : public IUnknown
    {
          HRESULT __stdcall CreateAggregatedInstance (
            BSTR bstrAssemblyName,
            BSTR bstrTypeName,
            IComAggregator* pOuterObject);
    };
}

