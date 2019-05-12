#include "stdafx.h"
#include "IfClsidHash.h"
#include "SmartHandle.h"
#include "SmartTYPEATTR.h"
#include "TypeLibInfo.h"
#include "CoClassInfo.h"
#include "VariantCollection.h"

using namespace ATL;

CIfClsidHash* CIfClsidHash::_This = NULL;

CIfClsidHash::CIfClsidHash()
{
}


CIfClsidHash::~CIfClsidHash()
{
}

HRESULT CIfClsidHash::createIndex()
{
    HRESULT hr;
    SmartHandle<HKEY, RegCloseKey> regKey;
    if (FAILED(hr = RegOpenKeyEx(HKEY_CLASSES_ROOT, L"TypeLib", 0, KEY_READ, &regKey))) {
        return hr;
    }
    clearCache();

    const LONG maxBufSize = 1024;
    wchar_t nameBuf[maxBufSize] = L"";
    for (LONG i = 0;; i++) {
        DWORD nameBufSize = maxBufSize;
        if (RegEnumKeyEx(regKey, i, nameBuf, &nameBufSize, 0, NULL, NULL, NULL) != ERROR_SUCCESS) {
            break;
        }
        CComBSTR tLibGuidStr(nameBuf);
        SmartHandle<HKEY, RegCloseKey> regKeyTlib;
        if (RegOpenKeyEx(HKEY_CLASSES_ROOT, str(L"TypeLib\\%s", tLibGuidStr).c_str(), 0, KEY_READ, &regKeyTlib) == ERROR_SUCCESS) {
            for (LONG j = 0;; j++) {
                DWORD nameBufSize = sizeof(nameBuf);
                if (RegEnumKeyEx(regKeyTlib, j, nameBuf, &nameBufSize, 0, NULL, NULL, NULL) != ERROR_SUCCESS) {
                    break;
                }
                CComBSTR versionStr(nameBuf);
                int major = 0, minor = 0;
                swscanf_s(nameBuf, L"%d.%d", &major, &minor);
                GUID tlibGuid;
                CLSIDFromString(tLibGuidStr, &tlibGuid);
                CComPtr<ITypeLib> typeLib;
                if (SUCCEEDED(LoadRegTypeLib(tlibGuid, (WORD)major, (WORD)minor, 0, &typeLib))) {
                    Values v;
                    v.tlibGuid = tlibGuid;
                    v.tlibMajor = major;
                    v.tlibMinor = minor;
                    UINT typeNum = typeLib->GetTypeInfoCount();
                    for (UINT k = 0; k < typeNum; k++) {
                        CComPtr<ITypeInfo> typeInfo;
                        typeLib->GetTypeInfo(k, &typeInfo);
                        SmartTYPEATTR typeAttr(typeInfo);
                        if (typeAttr.hasData() && typeAttr->typekind == TKIND_COCLASS) {
                            v.clsid = typeAttr->guid;
                            for (UINT l = 0; l < typeAttr->cImplTypes; l++) {
                                HREFTYPE refType;
                                if (SUCCEEDED(typeInfo->GetRefTypeOfImplType(l, &refType))) {
                                    CComPtr<ITypeInfo> ifTypeInfo;
                                    if (SUCCEEDED(typeInfo->GetRefTypeInfo(refType, &ifTypeInfo))) {
                                        SmartTYPEATTR ifTypeAttr(ifTypeInfo);
                                        int pos;
                                        if ((pos = _index.FindKey(ifTypeAttr->guid)) == -1) {
                                            CSimpleArray<Values> arr = CSimpleArray<Values>();
                                            arr.Add(v);
                                            _index.Add(ifTypeAttr->guid, arr);
                                        } else {
                                            _index.GetValueAt(pos).Add(v);
                                        }
                                    }
                                }
                            }
                        }
                    }
                };
            }
        }
    }
    return S_OK;
}

void CIfClsidHash::clearCache()
{
    _index.RemoveAll();
    _cache.RemoveAll();
}

CIfClsidHash& CIfClsidHash::entity()
{
    if (!_This) {
        _This = new CIfClsidHash();
        _This->createIndex();
    }
    return *_This;
}

HRESULT CIfClsidHash::getCoClass(GUID iid, CComVariant& result)
{
    HRESULT hr;
    if (!indexExists()) {
        if (FAILED((hr = createIndex()))) {
            return hr;
        }
    }
    int pos;
    if ((pos = _cache.FindKey(iid)) == -1) {
        int indexPos;
        if ((indexPos = _index.FindKey(iid)) == -1) {
            result.vt = VT_NULL;
            return S_FALSE;
        }
        CComPtr<IVariantCollection> variantCollection;
        CVariantCollection::CreateInstance(&variantCollection);
        CSimpleArray<Values>& v = _index.GetValueAt(indexPos);
        for (int i = 0; i < v.GetSize(); i++) {
            CComObject<CTypeLibInfo>* typeLibInfo;
            CComObject<CTypeLibInfo>::CreateInstance(&typeLibInfo);
            typeLibInfo->setup(v[i].tlibGuid, v[i].tlibMajor, v[i].tlibMinor);
            CComObject<CCoClassInfo>* coClassInfo;
            CComObject<CCoClassInfo>::CreateInstance(&coClassInfo);
            CComPtr<ITypeLibInfo> iTypeLibInfo;
            typeLibInfo->QueryInterface(&iTypeLibInfo);
            coClassInfo->setup(v[i].clsid, iTypeLibInfo);
            CComPtr<IDispatch> dispCoClassInfo;
            coClassInfo->QueryInterface(&dispCoClassInfo);
            CComVariant varCoClassInfo(dispCoClassInfo);
            variantCollection->Add(varCoClassInfo);
        }
        _cache.Add(iid, CComVariant(variantCollection));
        pos = _cache.FindKey(iid);
    }
    result.Copy(&_cache.GetValueAt(pos));
    return S_OK;
}
