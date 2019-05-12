#pragma once
#include "resource.h"
#include "ComUtility_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Not supported"
#endif

using namespace ATL;


// CPropertyDesc

class ATL_NO_VTABLE CPropertyDesc :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CPropertyDesc, &CLSID_PropertyDesc>,
    public ISupportErrorInfo,
    public IDispatchImpl<IPropertyDesc, &IID_IPropertyDesc, &LIBID_ComUtilityLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
protected:
    SHORT _vt;
    SHORT _invkind;
    CComBSTR _name;
    CComVariant _value;
    VARIANT_BOOL _isCollection;
    VARIANT_BOOL _valueUseful;
public:
    CPropertyDesc()
    {
        _isCollection = FALSE;
        _vt = 0;
        _invkind = 0;
        _valueUseful = FALSE;
    }

DECLARE_REGISTRY_RESOURCEID(IDR_PROPERTYDESC)


BEGIN_COM_MAP(CPropertyDesc)
    COM_INTERFACE_ENTRY(IPropertyDesc)
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
    STDMETHOD(put_Name)(BSTR newVal);
    STDMETHOD(get_Vt)(SHORT* pVal);
    STDMETHOD(put_Vt)(SHORT newVal);
    STDMETHOD(get_Invkind)(SHORT* pVal);
    STDMETHOD(put_Invkind)(SHORT newVal);
    STDMETHOD(get_IsCollection)(VARIANT_BOOL* pVal);
    STDMETHOD(put_IsCollection)(VARIANT_BOOL newVal);
    STDMETHOD(get_Value)(VARIANT* pVal);
    STDMETHOD(put_Value)(VARIANT newVal);
    STDMETHOD(get_ValueUseful)(VARIANT_BOOL* pVal);
    STDMETHOD(put_ValueUseful)(VARIANT_BOOL newVal);
};

OBJECT_ENTRY_AUTO(__uuidof(PropertyDesc), CPropertyDesc)
