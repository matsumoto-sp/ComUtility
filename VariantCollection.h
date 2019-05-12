#pragma once
#include "resource.h"
#include "ComUtility_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Not Supported"
#endif

using namespace ATL;


// CVariantCollection

class ATL_NO_VTABLE CVariantCollection :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CVariantCollection, &CLSID_VariantCollection>,
    public ISupportErrorInfo,
    public IDispatchImpl<IVariantCollection, &IID_IVariantCollection, &LIBID_ComUtilityLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
protected:
    std::vector<CComVariant> _values;
public:
    CVariantCollection()
    {
    }

DECLARE_REGISTRY_RESOURCEID(IDR_VariantCollection)


BEGIN_COM_MAP(CVariantCollection)
    COM_INTERFACE_ENTRY(IVariantCollection)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


    DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct()
    {
        return S_OK;
    }

    void FinalRelease()
    {
    }

public:



    STDMETHOD(get_Item)(LONG index, VARIANT* pVal);
    STDMETHOD(get__NewEnum)(IUnknown** pVal);
    STDMETHOD(get_Count)(LONG* pVal);
    STDMETHOD(Add)(VARIANT val);
};

OBJECT_ENTRY_AUTO(__uuidof(VariantCollection), CVariantCollection)
