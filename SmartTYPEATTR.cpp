#include "stdafx.h"
#include "SmartTYPEATTR.h"


SmartTYPEATTR::SmartTYPEATTR()
{

}

SmartTYPEATTR::SmartTYPEATTR(ITypeInfo* typeInfo)
{
    _typeInfo = typeInfo;
    _typeInfo->GetTypeAttr(&_data);

}

SmartTYPEATTR::~SmartTYPEATTR()
{
    if (_typeInfo) {
        _typeInfo->ReleaseTypeAttr(_data);
    }
}

