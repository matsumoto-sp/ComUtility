#include "stdafx.h"
#include "SmartVARDESC.h"


SmartVARDESC::SmartVARDESC()
{
}

SmartVARDESC::SmartVARDESC(ITypeInfo* typeInfo, UINT index)
{
    _typeInfo = typeInfo;
    _typeInfo->GetVarDesc(index, &_data);

}


SmartVARDESC::~SmartVARDESC()
{
    _typeInfo->ReleaseVarDesc(_data);
}
