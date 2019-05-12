#pragma once
#include "stdafx.h"
using namespace ATL;

class SmartTYPEATTR
{
protected:
    TYPEATTR* _data;
    CComPtr<ITypeInfo> _typeInfo;
public:
    SmartTYPEATTR();
    SmartTYPEATTR(ITypeInfo* typeInfo);
    virtual ~SmartTYPEATTR();
    bool hasData() { return _data != NULL; }
    TYPEATTR* operator -> () {
        return _data;
    }
};
