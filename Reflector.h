#pragma once
#include "resource.h"
#include "ComUtility_i.h"
#include "PropertyDesc.h"
#include "SmartFUNCDESC.h"
#include "SmartVARDESC.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Not supported"
#endif

using namespace ATL;


// CReflector

class ATL_NO_VTABLE CReflector :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CReflector, &CLSID_Reflector>,
    public ISupportErrorInfo,
    public IDispatchImpl<IReflector, &IID_IReflector, &LIBID_ComUtilityLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
protected:
    bool innerIsCollection(IDispatch* obj, SmartFUNCDESC& funcDesc);
    CComVariant getValue(IDispatch* obj, SmartFUNCDESC& funcDesc, CComPtr<IPropertyDesc> prop);
    CComVariant getValue(IDispatch* obj, SmartVARDESC& funcDesc, CComPtr<IPropertyDesc> prop);
public:
    CReflector()
    {
    }

DECLARE_REGISTRY_RESOURCEID(IDR_REFLECTOR)


BEGIN_COM_MAP(CReflector)
    COM_INTERFACE_ENTRY(IReflector)
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
    STDMETHOD(ScanIf)(IDispatch* obj, VARIANT* result);
    STDMETHOD(ScanDispIf)(IDispatch* obj, VARIANT* result);
    STDMETHOD(IsCollection)(IDispatch* obj, VARIANT_BOOL* result);
    STDMETHOD(SearchCoClass)(IDispatch* obj, VARIANT* result);
};

OBJECT_ENTRY_AUTO(__uuidof(Reflector), CReflector)
