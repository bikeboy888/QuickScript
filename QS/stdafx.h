// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#pragma comment(linker, "/nodefaultlib:libc.lib")
#pragma comment(linker, "/nodefaultlib:libcd.lib")

#ifndef STRICT
#define STRICT
#endif

#define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA

#include <ceconfig.h>
#if defined(WIN32_PLATFORM_PSPC) || defined(WIN32_PLATFORM_WFSP)
#define SHELL_AYGSHELL
#endif

// NOTE - this is value is not strongly correlated to the Windows CE OS version being targeted
#define WINVER _WIN32_WCE

#ifdef _CE_DCOM
#define _ATL_APARTMENT_THREADED
#endif
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit


#ifdef POCKETPC2003_UI_MODEL
#include "resourceppc.h"
#endif 

#include <atlbase.h>
#include <atlcom.h>
#include <atlsafe.h>
#include <atlctl.h>
#include <atlcoll.h>

#if defined(WIN32_PLATFORM_PSPC) || defined(WIN32_PLATFORM_WFSP)
#ifndef _DEVICE_RESOLUTION_AWARE
#define _DEVICE_RESOLUTION_AWARE
#endif
#endif

#ifdef _DEVICE_RESOLUTION_AWARE
#include "DeviceResolutionAware.h"
#endif

using namespace ATL;

#ifdef SHELL_AYGSHELL
#include <aygshell.h>
#pragma comment(lib, "aygshell.lib") 
#endif // SHELL_AYGSHELL

#if _WIN32_WCE < 0x500 && ( defined(WIN32_PLATFORM_PSPC) || defined(WIN32_PLATFORM_WFSP) )
	#pragma comment(lib, "ccrtrtti.lib")
	#ifdef _X86_	
		#if defined(_DEBUG)
			#pragma comment(lib, "libcmtx86d.lib")
		#else
			#pragma comment(lib, "libcmtx86.lib")
		#endif
	#endif
#endif

#include <altcecrt.h>

#include <activscp.h>
#include <wininet.h>
#pragma comment(lib, "wininet.lib")

#ifndef CHECKHR
#define CHECKHR(stmt) \
	hr = (stmt); \
	if (FAILED(hr)) \
	{ \
		return hr; \
	}
#endif
