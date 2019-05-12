#include "stdafx.h"
#include "Reflector.h"
#include "IfInfo.h"
#include "VariantCollection.h"
#include "RegistoryUtility.h"
#include "SmartTYPEATTR.h"
#include "SmartFUNCDESC.h"
#include "SmartVARDESC.h"
#include "SmartHandle.h"
#include "DispIfInfo.h"
#include "PropertyDesc.h"
#include "CoClassInfo.h"
#include "TypeLibInfo.h"
#include "IfClsidHash.h"


// CReflector

STDMETHODIMP CReflector::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* const arr[] = 
    {
        &IID_IReflector
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}

STDMETHODIMP CReflector::ScanIf(IDispatch* obj, VARIANT* result)
{
    SmartHandle<HKEY, RegCloseKey> regKey;
    if (RegOpenKeyEx(HKEY_CLASSES_ROOT, L"Interface", 0, KEY_READ, &regKey) != ERROR_SUCCESS) {
        return Error(L"Could not access registory 'HKCR/Interface'.");
    }
    CComPtr<IVariantCollection> variantCollection;
    CVariantCollection::CreateInstance(&variantCollection);
    for (LONG i = 0;; i++) {
        const LONG maxBufSize = 1024;
        wchar_t nameBuf[maxBufSize] = L"";
        DWORD nameBufSize = maxBufSize;
        if (RegEnumKeyEx(regKey, i, nameBuf, &nameBufSize, 0, NULL, NULL, NULL) != ERROR_SUCCESS) {
            break;
        }
        IID iid;
        CLSIDFromString(nameBuf, &iid);
        CComPtr<IUnknown> unk = NULL;
        if (SUCCEEDED(obj->QueryInterface(iid, (void**)(&unk)))) {
            CComPtr<IIfInfo> ifInfo;
            CIfInfo::CreateInstance(&ifInfo);
            ifInfo->SetupIf(nameBuf);
            variantCollection->Add(CComVariant(ifInfo));
        }
    }
    CComVariant variantVariantCollection = variantCollection.Detach();
    *result = variantVariantCollection;
    return S_OK;
}

bool CReflector::innerIsCollection(IDispatch* obj, SmartFUNCDESC& funcDesc)
{
    if (funcDesc->elemdescFunc.tdesc.vt == VT_DISPATCH || funcDesc->elemdescFunc.tdesc.vt == VT_PTR) {
        if (funcDesc->invkind == INVOKE_FUNC || funcDesc->invkind == INVOKE_PROPERTYGET) {
            CComVariant r;
            DISPPARAMS params;
            params.cArgs = 0;
            params.cNamedArgs = 0;

            if (SUCCEEDED(obj->Invoke(funcDesc->memid, IID_NULL, NULL, funcDesc->invkind, &params, &r, NULL, NULL))) {
                if (r.vt == VT_DISPATCH && r.pdispVal) {
                    CComPtr<ITypeInfo> typeInfo;
                    if (SUCCEEDED(r.pdispVal->GetTypeInfo(0, 0, &typeInfo))) {
                        SmartTYPEATTR typeAttr(typeInfo);
                        for (WORD i = 0; i < typeAttr->cFuncs; i++) {
                            SmartFUNCDESC funcDesc(typeInfo, i);
                            if (funcDesc->memid == DISPID_NEWENUM) {
                                return true;
                            }
                        }
                    }
                }
            }
        }
    }
    return false;
}

CComVariant CReflector::getValue(IDispatch* obj, SmartFUNCDESC& funcDesc, CComPtr<IPropertyDesc> prop)
{
    if (funcDesc->invkind == INVOKE_FUNC || funcDesc->invkind == INVOKE_PROPERTYGET) {
        CComVariant r;
        DISPPARAMS params;
        params.cArgs = 0;
        params.cNamedArgs = 0;
        if (SUCCEEDED(obj->Invoke(funcDesc->memid, IID_NULL, NULL, funcDesc->invkind, &params, &r, NULL, NULL))) {
            prop->put_ValueUseful(TRUE);
            prop->put_Value(r);
            return r;
        }
    }
    prop->put_ValueUseful(FALSE);
    return false;
}

CComVariant CReflector::getValue(IDispatch* obj, SmartVARDESC& funcDesc, CComPtr<IPropertyDesc> prop)
{
    CComVariant r;
    DISPPARAMS params;
    params.cArgs = 0;
    params.cNamedArgs = 0;
    if (SUCCEEDED(obj->Invoke(funcDesc->memid, IID_NULL, NULL, INVOKE_PROPERTYGET, &params, &r, NULL, NULL))) {
        prop->put_ValueUseful(TRUE);
        prop->put_Value(r);
        return r;
    }
    prop->put_ValueUseful(FALSE);
    return false;
}


STDMETHODIMP CReflector::ScanDispIf(IDispatch* obj, VARIANT* result)
{
    if (!obj) {
        return Error(L"Null Pointer exception");
    }
    CComPtr<ITypeInfo> typeInfo;
    HRESULT hr;
    if (SUCCEEDED(obj->GetTypeInfo(0, 0, &typeInfo))) {
        SmartTYPEATTR typeAttr(typeInfo);
        CComPtr<IDispIfInfo> dispIfInfo;
        CDispIfInfo::CreateInstance(&dispIfInfo);
        CSimpleMap<CComBSTR, bool> propertiesCheck;
        CSimpleMap<CComBSTR, bool> properties;
        CComPtr<IDispatch> dispFunctions;
        CComPtr<IVariantCollection> functions;
        dispIfInfo->get_Functions(&dispFunctions);
        dispFunctions->QueryInterface(&functions);
        CComPtr<IDispatch> dispProperties;
        CComPtr<IVariantCollection> resultProperties;
        dispIfInfo->get_Properties(&dispProperties);
        dispProperties->QueryInterface(&resultProperties);
        for (WORD i = 0; i < typeAttr->cFuncs; i++) {
            SmartFUNCDESC funcDesc(typeInfo, i);
            const UINT maxNames = 300;
            BSTR names[maxNames] = {};
            UINT namesNum = 0;
            if (FAILED(hr = typeInfo->GetNames(funcDesc->memid, names, maxNames, &namesNum))) {
                return hr;
            }
            CComPtr<IPropertyDesc> prop;
            CPropertyDesc::CreateInstance(&prop);
            prop->put_Name(names[0]);
            prop->put_Vt(funcDesc->elemdescFunc.tdesc.vt);
            prop->put_Invkind(funcDesc->invkind);
            functions->Add(CComVariant(prop));
            if (propertiesCheck.FindKey(names[0]) != -1) {
                properties.Add(names[0], true);
            }
            propertiesCheck.Add(names[0], true);
            for (UINT i = 0; i < namesNum; i++) {
                SysFreeString(names[i]);
            }
        }
        for (WORD i = 0; i < typeAttr->cFuncs; i++) {
            SmartFUNCDESC funcDesc(typeInfo, i);
            const UINT maxNames = 300;
            BSTR names[maxNames] = {};
            UINT namesNum = 0;
            if (FAILED(hr = typeInfo->GetNames(funcDesc->memid, names, maxNames, &namesNum))) {
                return hr;
            }
            bool isCol = innerIsCollection(obj, funcDesc);
            if ((funcDesc->invkind == INVOKE_PROPERTYGET && properties.FindKey(names[0]) != -1) || isCol) {
                CComPtr<IPropertyDesc> prop;
                CPropertyDesc::CreateInstance(&prop);
                prop->put_Name(names[0]);
                prop->put_Vt(funcDesc->elemdescFunc.tdesc.vt);
                prop->put_Invkind(funcDesc->invkind);
                prop->put_IsCollection(isCol ? TRUE : FALSE);
                getValue(obj, funcDesc, prop);
                resultProperties->Add(CComVariant(prop));
            }
        }
        for (WORD i = 0; i < typeAttr->cVars; i++) {
            SmartVARDESC varDesc(typeInfo, i);
            const UINT maxNames = 300;
            BSTR names[maxNames] = {};
            UINT namesNum = 0;
            if (FAILED(hr = typeInfo->GetNames(varDesc->memid, names, maxNames, &namesNum))) {
                return hr;
            }
            CComPtr<IPropertyDesc> prop;
            CPropertyDesc::CreateInstance(&prop);
            prop->put_Name(names[0]);
            prop->put_Vt(varDesc->elemdescVar.tdesc.vt);
            prop->put_Invkind(INVOKE_PROPERTYGET);
            prop->put_IsCollection(FALSE);
            getValue(obj, varDesc, prop);
            resultProperties->Add(CComVariant(prop));
        }
        CComVariant out = dispIfInfo.Detach();
        *result = out;
        return S_OK;
    }
    CComVariant out(NULL);
    out.Detach(result);
    return S_FALSE;
}


STDMETHODIMP CReflector::IsCollection(IDispatch* obj, VARIANT_BOOL* result)
{
    CComPtr<ITypeInfo> typeInfo;
    HRESULT hr;
    if (FAILED(hr = obj->GetTypeInfo(0, 0, &typeInfo))) {
        return hr;
    }
    *result = FALSE;
    SmartTYPEATTR typeAttr(typeInfo);
    for (WORD i = 0; i < typeAttr->cFuncs; i++) {
        SmartFUNCDESC funcDesc(typeInfo, i);
        if (funcDesc->memid == DISPID_NEWENUM) {
            *result = TRUE;
            return S_OK;
        }
    }
    return S_OK;
}


STDMETHODIMP CReflector::SearchCoClass(IDispatch* obj, VARIANT* result)
{
    if (!obj) {
        return Error(L"Null Pointer exception");
    }
    CComPtr<ITypeInfo> typeInfo;
    HRESULT hr;
    if (FAILED(hr = obj->GetTypeInfo(0, 0, &typeInfo))) {
        return hr;
    }
    SmartTYPEATTR typeAttr(typeInfo);
    CComVariant out;
    if (FAILED(hr = CIfClsidHash::entity().getCoClass(typeAttr->guid, out))) {
        return hr;
    }
    out.Detach(result);
    return S_OK;
}

