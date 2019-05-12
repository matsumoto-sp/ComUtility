#include "stdafx.h"
#include "SmartTLIBATTR.h"


SmartTLIBATTR::SmartTLIBATTR()
{

}

SmartTLIBATTR::SmartTLIBATTR(ITypeLib* typeLib)
{
    _typeLib = typeLib;
    _typeLib->GetLibAttr(&_data);

}

SmartTLIBATTR::~SmartTLIBATTR()
{
    if (_typeLib) {
        _typeLib->ReleaseTLibAttr(_data);
    }
}

