#include "stdafx.h"
#include "CoClassInfo.h"
#include "RegistoryUtility.h"

// CCoClassInfo

STDMETHODIMP CCoClassInfo::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* const arr[] = 
    {
        &IID_ICoClassInfo
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}


HRESULT CCoClassInfo::setup(GUID coClassGUID, ITypeLibInfo* typeLibInfo)
{
    RegistoryUtility ru(L"CLSID", coClassGUID);
    _CLSID = coClassGUID;
    _InprocServer32 = ru.nodeValue(L"InprocServer32");
    _ProgID = ru.nodeValue(L"ProgID");
    _TypeLib = typeLibInfo;
    return S_OK;
}


STDMETHODIMP CCoClassInfo::get_CLSID(BSTR* pVal)
{
    *pVal = _CLSID.Copy();
    return S_OK;
}


STDMETHODIMP CCoClassInfo::get_InprocServer32(BSTR* pVal)
{
    *pVal = _InprocServer32.Copy();
    return S_OK;
}


STDMETHODIMP CCoClassInfo::get_ProgID(BSTR* pVal)
{
    *pVal = _ProgID.Copy();
    return S_OK;
}


STDMETHODIMP CCoClassInfo::get_TypeLibInfo(IDispatch** pVal)
{
    return _TypeLib.QueryInterface(pVal);
}
