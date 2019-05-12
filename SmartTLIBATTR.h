#pragma once
#include "stdafx.h"
using namespace ATL;

class SmartTLIBATTR
{
protected:
    TLIBATTR* _data;
    CComPtr<ITypeLib> _typeLib;
public:
    SmartTLIBATTR();
    SmartTLIBATTR(ITypeLib* typeLib);
    virtual ~SmartTLIBATTR();
    TLIBATTR* operator -> () {
        return _data;
    }
};