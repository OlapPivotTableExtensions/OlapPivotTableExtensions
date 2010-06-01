// stdafx.h 
#pragma once

#ifndef STRICT
#define STRICT
#endif

#ifndef WINVER				// Allow use of features specific to Windows XP or later.
#define WINVER 0x0501		// Change this to the appropriate value to target other versions of Windows.
#endif

#ifndef _WIN32_WINNT		// Allow use of features specific to Windows XP or later.                   
#define _WIN32_WINNT 0x0501	// Change this to the appropriate value to target other versions of Windows.
#endif						

#ifndef _WIN32_WINDOWS		// Allow use of features specific to Windows 98 or later.
#define _WIN32_WINDOWS 0x0410 // Change this to the appropriate value to target Windows Me or later.
#endif

#ifndef _WIN32_IE			// Allow use of features specific to IE 6.0 or later.
#define _WIN32_IE 0x0600	// Change this to the appropriate value to target other versions of IE.
#endif

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// Some CString constructors will be explicit.
#define _ATL_ALL_WARNINGS	// Turns off ATL's hiding of some safely ignored warning messages.

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>

#include "interop.h"

#pragma warning( disable : 4278 )
#pragma warning( disable : 4146 )
    // For _AppDomain. Used to communicate with the default app domain from unmanaged code.
    #import <mscorlib.tlb> raw_interfaces_only high_property_prefixes("_get","_put","_putref")

    // Imports the MSADDNDR.DLL typelib which we need for IDTExtensibility2.
    #import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4" raw_interfaces_only named_guids
#pragma warning( default : 4146 )
#pragma warning( default : 4278 )

using namespace ATL;

// For CorBindToRuntimeEx and ICorRuntimeHost.
#include <mscoree.h>

#define IfFailGo(x) { hr=(x); if (FAILED(hr)) goto Error; }
#define IfNullGo(p) { if(!p) {hr = E_FAIL; goto Error; } }

#include <windows.h>
#include <assert.h>

// Additional statements for Aggregator.
#include <new>
#include <strsafe.h>
