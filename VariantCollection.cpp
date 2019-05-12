#include "stdafx.h"
#include "VariantCollection.h"


// CVariantCollection

STDMETHODIMP CVariantCollection::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* const arr[] = 
    {
        &IID_IVariantCollection
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}


STDMETHODIMP CVariantCollection::get_Item(LONG index, VARIANT* pVal)
{
    if (index < 0 || (LONG)_values.size() <= index) {
        return Error(L"Out of range.");
    }
    CComVariant out(_values[index]);
    out.Detach(pVal);
    return S_OK;
}


STDMETHODIMP CVariantCollection::get__NewEnum(IUnknown** pVal)
{
    CComObject<CComEnum<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT> > >* enumlator
        = new CComObject<CComEnum<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT> > >();
    if (_values.size()) {
        enumlator->Init(&_values[0], &_values[0] + _values.size(), NULL, AtlFlagCopy);
    }
    enumlator->QueryInterface(IID_IUnknown, (void**)pVal);
    return S_OK;
}


STDMETHODIMP CVariantCollection::get_Count(LONG* pVal)
{
    *pVal = (LONG)_values.size();
    return S_OK;
}


STDMETHODIMP CVariantCollection::Add(VARIANT val)
{
    _values.push_back(val);
    return S_OK;
}
