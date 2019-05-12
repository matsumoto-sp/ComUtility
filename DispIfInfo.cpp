#include "stdafx.h"
#include "DispIfInfo.h"
#include "RegistoryUtility.h"

// CDispIfInfo

STDMETHODIMP CDispIfInfo::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* const arr[] = 
    {
        &IID_IDispIfInfo
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}


STDMETHODIMP CDispIfInfo::get_Functions(IDispatch** pVal)
{
    return _Functions.QueryInterface(pVal);
}


STDMETHODIMP CDispIfInfo::get_Properties(IDispatch** pVal)
{
    return _Properties.QueryInterface(pVal);
}
