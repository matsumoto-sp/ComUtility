#include "stdafx.h"
#include "resource.h"
#include "ComUtility_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CComUtilityModule _AtlModule;

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
#ifdef _MERGE_PROXYSTUB
    if (!PrxDllMain(hInstance, dwReason, lpReserved))
        return FALSE;
#endif
    hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved); 
}
