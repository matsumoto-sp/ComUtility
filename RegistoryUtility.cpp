#include "stdafx.h"
#include "RegistoryUtility.h"
#include "SmartHandle.h"

RegistoryUtility::RegistoryUtility(BSTR trunk, BSTR iid)
{
    _trunk = trunk;
    _iid = iid;
}

RegistoryUtility::RegistoryUtility(BSTR trunk, GUID iid)
{
    _trunk = trunk;
    wchar_t buf[100];
    StringFromGUID2(iid, buf, 100);
    _iid = buf;
}

RegistoryUtility::~RegistoryUtility()
{
}

CComBSTR RegistoryUtility::getKeyValue(BSTR key, BSTR value)
{
    SmartHandle<HKEY, RegCloseKey> hKey;
    CComBSTR fullKey = str(L"%s\\%s", _trunk, _iid).c_str();
    if (key) {
        fullKey += "\\";
        fullKey += key;
    }
    if (RegOpenKeyEx(HKEY_CLASSES_ROOT, fullKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        const int maxBuffSize = 1000;
        DWORD buffSize = maxBuffSize * sizeof(wchar_t);
        wchar_t buf[maxBuffSize];
        if (RegQueryValueEx(hKey, value ? value : L"", NULL, NULL, (LPBYTE)buf, &buffSize) == ERROR_SUCCESS) {
            return buf;
        }
    }
    return NULL;
}

CComBSTR RegistoryUtility::value()
{
    return getKeyValue(NULL, NULL);
}
CComBSTR RegistoryUtility::nodeValue(BSTR node)
{
    return getKeyValue(node, NULL);
}

