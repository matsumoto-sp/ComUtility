#pragma once
#include "stdafx.h"
using namespace ATL;

class SmartVARDESC
{
protected:
    VARDESC* _data;
    CComPtr<ITypeInfo> _typeInfo;
public:
    SmartVARDESC();
    SmartVARDESC(ITypeInfo* typeInfo, UINT index);
    virtual ~SmartVARDESC();
    VARDESC* operator -> () {
        return _data;
    }
};

