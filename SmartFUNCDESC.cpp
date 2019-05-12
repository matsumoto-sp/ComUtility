#include "stdafx.h"
#include "SmartFUNCDESC.h"


SmartFUNCDESC::SmartFUNCDESC()
{
}

SmartFUNCDESC::SmartFUNCDESC(ITypeInfo* typeInfo, UINT index)
{
    _typeInfo = typeInfo;
    _typeInfo->GetFuncDesc(index, &_data);

}


SmartFUNCDESC::~SmartFUNCDESC()
{
    _typeInfo->ReleaseFuncDesc(_data);
}
