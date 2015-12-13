// IfsToolbox.cpp : Implementation of DLL Exports.


// Note: Proxy/Stub Information
//      To merge the proxy/stub code into the object DLL, add the file 
//      dlldatax.c to the project.  Make sure precompiled headers 
//      are turned off for this file, and add _MERGE_PROXYSTUB to the 
//      defines for the project.  
//
//      If you are not running WinNT4.0 or Win95 with DCOM, then you
//      need to remove the following define from dlldatax.c
//      #define _WIN32_WINNT 0x0400
//
//      Further, if you are running MIDL without /Oicf switch, you also 
//      need to remove the following define from dlldatax.c.
//      #define USE_STUBLESS_PROXY
//
//      Modify the custom build rule for IfsToolbox.idl by adding the following 
//      files to the Outputs.
//          IfsToolbox_p.c
//          dlldata.c
//      To build a separate proxy/stub DLL, 
//      run nmake -f IfsToolboxps.mk in the project directory.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include "IfsToolbox.h"
#include "dlldatax.h"

#include "IfsToolbox_i.c"
#include "Toolbox.h"
#include "ToolBoxTabs.h"
#include "ToolBoxTab.h"
#include "ToolBoxItem.h"
#include "ToolBoxItems.h"
#include "PPToolbox.h"

#ifdef _MERGE_PROXYSTUB
extern "C" HINSTANCE hProxyDll;
#endif

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_Toolbox, CToolbox)
OBJECT_ENTRY(CLSID_ToolBoxTabs, CToolBoxTabs)
OBJECT_ENTRY(CLSID_ToolBoxTab, CToolBoxTab)
OBJECT_ENTRY(CLSID_ToolBoxItem, CToolBoxItem)
OBJECT_ENTRY(CLSID_ToolBoxItems, CToolBoxItems)
OBJECT_ENTRY(CLSID_PPToolbox, CPPToolbox)
END_OBJECT_MAP()

/////////////////////////////////////////////////////////////////////////////
// DLL Entry Point

extern "C"
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    lpReserved;
#ifdef _MERGE_PROXYSTUB
    if (!PrxDllMain(hInstance, dwReason, lpReserved))
        return FALSE;
#endif
    if (dwReason == DLL_PROCESS_ATTACH)
    {
		//Initialize Common controls that we may use ICC_WIN95_CLASSES
		UINT nFlags = ICC_WIN95_CLASSES |ICC_USEREX_CLASSES |ICC_COOL_CLASSES | ICC_HOTKEY_CLASS|ICC_WIN95_CLASSES|
			ICC_BAR_CLASSES | ICC_TAB_CLASSES |ICC_PROGRESS_CLASS | ICC_UPDOWN_CLASS | ICC_TREEVIEW_CLASSES|ICC_PAGESCROLLER_CLASS;
		AtlInitCommonControls(nFlags);   // add flags to support other controls
		
        _Module.Init(ObjectMap, hInstance, &LIBID_IFSTOOLBOXLib);
        DisableThreadLibraryCalls(hInstance);
    }
    else if (dwReason == DLL_PROCESS_DETACH)
        _Module.Term();
    return TRUE;    // ok
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllCanUnloadNow() != S_OK)
        return S_FALSE;
#endif
    return (_Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllGetClassObject(rclsid, riid, ppv) == S_OK)
        return S_OK;
#endif
    return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
#ifdef _MERGE_PROXYSTUB
    HRESULT hRes = PrxDllRegisterServer();
    if (FAILED(hRes))
        return hRes;
#endif
    // registers object, typelib and all interfaces in typelib
    return _Module.RegisterServer(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
#ifdef _MERGE_PROXYSTUB
    PrxDllUnregisterServer();
#endif
    return _Module.UnregisterServer(TRUE);
}


