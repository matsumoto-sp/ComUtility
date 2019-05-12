#include "stdafx.h"
#include "PropertyDesc.h"


// CPropertyDesc

STDMETHODIMP CPropertyDesc::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* const arr[] = 
    {
        &IID_IPropertyDesc
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}


STDMETHODIMP CPropertyDesc::get_Name(BSTR* pVal)
{
    *pVal = _name.Copy();
    return S_OK;
}


STDMETHODIMP CPropertyDesc::put_Name(BSTR newVal)
{
    _name = newVal;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::get_Vt(SHORT* pVal)
{
    *pVal = _vt;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::put_Vt(SHORT newVal)
{
    _vt = newVal;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::get_Invkind(SHORT* pVal)
{
    *pVal = _invkind;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::put_Invkind(SHORT newVal)
{
    _invkind = newVal;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::get_IsCollection(VARIANT_BOOL* pVal)
{
    *pVal = _isCollection;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::put_IsCollection(VARIANT_BOOL newVal)
{
    _isCollection = newVal;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::get_Value(VARIANT* pVal)
{
    CComVariant out = _value;
    out.Detach(pVal);
    return S_OK;
}


STDMETHODIMP CPropertyDesc::put_Value(VARIANT newVal)
{
    _value.Copy(&newVal);
    return S_OK;
}


STDMETHODIMP CPropertyDesc::get_ValueUseful(VARIANT_BOOL* pVal)
{
    VARIANT_BOOL out = _valueUseful;
    if (out && _value.vt == VT_DISPATCH && !_value.pdispVal) {
        out = FALSE;
    }
    *pVal = out;
    return S_OK;
}


STDMETHODIMP CPropertyDesc::put_ValueUseful(VARIANT_BOOL newVal)
{
    _valueUseful = newVal;
    return S_OK;
}
