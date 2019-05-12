#pragma once
#include "stdafx.h"

using namespace ATL;

class RegistoryUtility
{
protected:
    CComBSTR _iid;
    CComBSTR _trunk;
    CComBSTR getKeyValue(BSTR key, BSTR value);
public:
    RegistoryUtility(BSTR trunk, BSTR iid);
    RegistoryUtility(BSTR trunk, GUID iid);
    ~RegistoryUtility();
    CComBSTR value();
    CComBSTR nodeValue(BSTR node);
};

