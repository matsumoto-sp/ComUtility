#pragma once
#include "resource.h"
#include "ComUtility_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Not supported"
#endif

using namespace ATL;


// CIfInfo

class ATL_NO_VTABLE CIfInfo :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CIfInfo, &CLSID_IfInfo>,
    public ISupportErrorInfo,
    public IDispatchImpl<IIfInfo, &IID_IIfInfo, &LIBID_ComUtilityLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
protected:
    CComBSTR _name;
    CComBSTR _iid;
    CComBSTR _typeLib;
public:
    CIfInfo()
    {
    }

DECLARE_REGISTRY_RESOURCEID(IDR_IFINFO)


BEGIN_COM_MAP(CIfInfo)
    COM_INTERFACE_ENTRY(IIfInfo)
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



    STDMETHOD(get_Name)(BSTR* pVal);
    STDMETHOD(get_IID)(BSTR* pVal);
    STDMETHOD(get_TypeLib)(BSTR* pVal);
    STDMETHOD(SetupIf)(BSTR iid);
};

OBJECT_ENTRY_AUTO(__uuidof(IfInfo), CIfInfo)
