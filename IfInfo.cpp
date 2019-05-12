#include "stdafx.h"
#include "IfInfo.h"
#include "RegistoryUtility.h"


// CIfInfo
STDMETHODIMP CIfInfo::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* const arr[] =
    {
        &IID_IReflector
    };

    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i], riid))
            return S_OK;
    }
    return S_FALSE;
}


STDMETHODIMP CIfInfo::get_Name(BSTR* pVal)
{
    *pVal = nullOrBstr(_name);
    return S_OK;
}


STDMETHODIMP CIfInfo::get_IID(BSTR* pVal)
{
    *pVal = nullOrBstr(_iid);
    return S_OK;
}

STDMETHODIMP CIfInfo::get_TypeLib(BSTR* pVal)
{
    *pVal = nullOrBstr(_typeLib);
    return S_OK;
}

STDMETHODIMP CIfInfo::SetupIf(BSTR iid)
{
    RegistoryUtility ru(L"Interface", iid);
    _iid = iid;
    _name = ru.value();
    _typeLib = ru.nodeValue(L"TypeLib");
    return S_OK;
}
