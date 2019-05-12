#pragma once
#include "stdafx.h"
using namespace ATL;

class SmartFUNCDESC
{
protected:
    FUNCDESC* _data;
    CComPtr<ITypeInfo> _typeInfo;
public:
    SmartFUNCDESC();
    SmartFUNCDESC(ITypeInfo* typeInfo, UINT index);
    virtual ~SmartFUNCDESC();
    FUNCDESC* operator -> () {
        return _data;
    }
};

