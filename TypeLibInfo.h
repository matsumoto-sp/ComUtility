#pragma once
#include "resource.h"
#include "ComUtility_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Not supported"
#endif

using namespace ATL;


// CTypeLibInfo

class ATL_NO_VTABLE CTypeLibInfo :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CTypeLibInfo, &CLSID_TypeLibInfo>,
    public ISupportErrorInfo,
    public IDispatchImpl<ITypeLibInfo, &IID_ITypeLibInfo, &LIBID_ComUtilityLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
protected:
    CComBSTR _guid;
    CComBSTR _name;
    SHORT _verMajor;
    SHORT _verMinor;
    CComBSTR _typeLibFile32;
    CComBSTR _typeLibFile64;
public:
    CTypeLibInfo()
    {
    }
    HRESULT setup(BSTR guid, BSTR version);
    HRESULT setup(CComPtr<ITypeLib>& typeLib);
    HRESULT setup(GUID guid, SHORT major, SHORT minor);

DECLARE_REGISTRY_RESOURCEID(IDR_TYPELIBINFO)


BEGIN_COM_MAP(CTypeLibInfo)
    COM_INTERFACE_ENTRY(ITypeLibInfo)
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

    STDMETHOD(get_GUID)(BSTR* pVal);
    STDMETHOD(get_Name)(BSTR* pVal);
    STDMETHOD(get_VerMajor)(SHORT* pVal);
    STDMETHOD(get_VerMinor)(SHORT* pVal);
    STDMETHOD(get_TypeLibFile32)(BSTR* pVal);
    STDMETHOD(get_TypeLibFile64)(BSTR* pVal);
};

OBJECT_ENTRY_AUTO(__uuidof(TypeLibInfo), CTypeLibInfo)
