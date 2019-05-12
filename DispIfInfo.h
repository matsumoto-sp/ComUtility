#pragma once
#include "resource.h"
#include "ComUtility_i.h"
#include "VariantCollection.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Not supported"
#endif

using namespace ATL;


// CDispIfInfo

class ATL_NO_VTABLE CDispIfInfo :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CDispIfInfo, &CLSID_DispIfInfo>,
    public ISupportErrorInfo,
    public IDispatchImpl<IDispIfInfo, &IID_IDispIfInfo, &LIBID_ComUtilityLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
protected:
    CComPtr<IVariantCollection> _Functions;
    CComPtr<IVariantCollection> _Properties;

public:
    CDispIfInfo()
    {
        CVariantCollection::CreateInstance(&_Functions);
        CVariantCollection::CreateInstance(&_Properties);
    }

DECLARE_REGISTRY_RESOURCEID(IDR_DISPIFINFO)


BEGIN_COM_MAP(CDispIfInfo)
    COM_INTERFACE_ENTRY(IDispIfInfo)
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
    STDMETHOD(get_Functions)(IDispatch** pVal);
    STDMETHOD(get_Properties)(IDispatch** pVal);
};

OBJECT_ENTRY_AUTO(__uuidof(DispIfInfo), CDispIfInfo)
