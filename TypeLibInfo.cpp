#include "stdafx.h"
#include "TypeLibInfo.h"
#include "RegistoryUtility.h"
#include "SmartTLIBATTR.h"


// CTypeLibInfo

STDMETHODIMP CTypeLibInfo::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* const arr[] = 
    {
        &IID_ITypeLibInfo
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}

HRESULT CTypeLibInfo::setup(BSTR guid, BSTR version)
{
    _guid = guid;
    int major, minor;
    swscanf_s(version, L"%d.%d", &major, &minor);
    _verMajor = (SHORT)major;
    _verMinor = (SHORT)minor;
    CComBSTR key = str(L"%s\\%s", guid, version).c_str();
    RegistoryUtility ru(L"TypeLib", key);
    _name = ru.value();
    CComBSTR key0 = str(L"%s\\%s\\0", guid, version).c_str();
    RegistoryUtility ru0(L"TypeLib", key0);
    _typeLibFile32 = ru0.nodeValue(L"win32");
    _typeLibFile64 = ru0.nodeValue(L"win64");
    return S_OK;
}


HRESULT CTypeLibInfo::setup(GUID guid, SHORT major, SHORT minor)
{
    setup(CComBSTR(guid), CComBSTR(str(L"%d.%d", major, minor).c_str()));
    return S_OK;
}


HRESULT CTypeLibInfo::setup(CComPtr<ITypeLib>& typeLib)
{
    SmartTLIBATTR attr(typeLib);
    setup(attr->guid, attr->wMajorVerNum, attr->wMinorVerNum);
    return S_OK;
}


STDMETHODIMP CTypeLibInfo::get_GUID(BSTR* pVal)
{
    *pVal = _guid.Copy();
    return S_OK;
}


STDMETHODIMP CTypeLibInfo::get_Name(BSTR* pVal)
{
    *pVal = _name.Copy();
    return S_OK;
}


STDMETHODIMP CTypeLibInfo::get_VerMajor(SHORT* pVal)
{
    *pVal = _verMajor;
    return S_OK;
}


STDMETHODIMP CTypeLibInfo::get_VerMinor(SHORT* pVal)
{
    *pVal = _verMinor;
    return S_OK;
}


STDMETHODIMP CTypeLibInfo::get_TypeLibFile32(BSTR* pVal)
{
    *pVal = _typeLibFile32.Copy();
    return S_OK;
}


STDMETHODIMP CTypeLibInfo::get_TypeLibFile64(BSTR* pVal)
{
    *pVal = _typeLibFile64.Copy();
    return S_OK;
}
