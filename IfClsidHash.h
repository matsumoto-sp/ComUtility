#pragma once
#include "stdafx.h"
class CIfClsidHash
{
public:
    struct Values {
        GUID tlibGuid;
        SHORT tlibMajor;
        SHORT tlibMinor;
        CLSID clsid;
    };

protected:
    CIfClsidHash();
    static CIfClsidHash* _This;
    ATL::CSimpleMap<IID, ATL::CSimpleArray<Values> > _index;
    ATL::CSimpleMap<IID, ATL::CComVariant> _cache;
    bool indexExists() {
        return _index.GetSize() != 0;
    }
public:
    virtual ~CIfClsidHash();
    HRESULT createIndex();
    void clearCache();
    static CIfClsidHash& entity();
    HRESULT getCoClass(GUID iid, ATL::CComVariant& result);
};

