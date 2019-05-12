#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS


#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <atlsafe.h>
#include <string>
#include <stdarg.h>
#include <winreg.h>
#include <vector>
#include "rpc.h"
#include "rpcndr.h"
#include "windows.h"
#include "ole2.h"
#include "oaidl.h"
#include "ocidl.h"



inline std::wstring str(const wchar_t* format, ...)
{
    va_list arg;
    const int maxSize = 1024;
    wchar_t buf[maxSize];
    va_start(arg, format);
    vswprintf_s(buf, maxSize, format, arg);
    va_end(arg);
    return std::wstring(buf);
}

inline BSTR nullOrBstr(ATL::CComBSTR& bstr)
{
    return bstr.Length() ? bstr.Copy() : NULL;
}
