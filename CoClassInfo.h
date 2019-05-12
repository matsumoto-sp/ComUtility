// CoClassInfo.h : CCoClassInfo ÇÃêÈåæ

#pragma once
#include "resource.h"       // ÉÅÉCÉì ÉVÉìÉ{Éã



#include "ComUtility_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Not supported"
#endif

using namespace ATL;


// CCoClassInfo

class ATL_NO_VTABLE CCoClassInfo :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CCoClassInfo, &CLSID_CoClassInfo>,
    public ISupportErrorInfo,
    public IDispatchImpl<ICoClassInfo, &IID_ICoClassInfo, &LIBID_ComUtilityLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
protected:
    CComBSTR _CLSID;
    CComBSTR _InprocServer32;
    CComBSTR _ProgID;
    CComPtr<ITypeLibInfo> _TypeLib;
public:
    CCoClassInfo()
    {
    }
    HRESULT setup(GUID coClassGUID, ITypeLibInfo* typeLibInfo);

DECLARE_REGISTRY_RESOURCEID(IDR_COCLASSINFO)


BEGIN_COM_MAP(CCoClassInfo)
    COM_INTERFACE_ENTRY(ICoClassInfo)
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


    STDMETHOD(get_CLSID)(BSTR* pVal);
    STDMETHOD(get_InprocServer32)(BSTR* pVal);
    STDMETHOD(get_ProgID)(BSTR* pVal);
    STDMETHOD(get_TypeLibInfo)(IDispatch** pVal);
};

OBJECT_ENTRY_AUTO(__uuidof(CoClassInfo), CCoClassInfo)
