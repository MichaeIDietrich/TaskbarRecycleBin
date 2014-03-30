//#include <strsafe.h> // for StringCchXXX functions
//#include <olectl.h> // for SELFREG_E_CLASS
//#include <shlobj.h> // for ICatRegister

#include "StdAfx.h"
#include "DllMain.h"
#include "Guids.h"


HINSTANCE g_hInst = nullptr;


CRecycleBinDeskBandModule _AtlModule;

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    g_hInst = hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved);
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
    return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    return _AtlModule.DllRegisterServer(/*bRegTypeLib =*/ FALSE);
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
    return _AtlModule.DllUnregisterServer(/*bUnRegTypeLib =*/ FALSE);
}

// DllInstall - Adds/Removes entries to the system registry per user
//              per machine.	
STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
    HRESULT hr = E_FAIL;
    static const wchar_t szUserSwitch[] = _T("user");

    if (pszCmdLine != NULL)
    {
        if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
        {
            AtlSetPerUserRegistration(true);
        }
    }

    if (bInstall)
    {
        hr = DllRegisterServer();
        if (FAILED(hr))
        {
            DllUnregisterServer();
        }
    }
    else
    {
        hr = DllUnregisterServer();
    }

    return hr;
}