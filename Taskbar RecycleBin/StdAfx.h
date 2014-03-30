// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

//#include "targetver.h"

#define _CRT_SECURE_CPP_OVERLOAD_STANDARD_NAMES 1   // Enable template overloads of standard CRT functions
#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers
#define VC_EXTRALEAN

// Windows header files
#include <Windows.h>
#include <WindowsX.h>
#include <ShlObj.h>
#include <shellapi.h>

// additional string header files
#include <tchar.h>
#include <strsafe.h> // for StringCchXXX functions

// GDI+
#include <gdiplus.h>
#pragma comment(lib, "gdiplus.lib")

// Visual styles
#include <UxTheme.h>
#include <VsSym32.h>
#pragma comment(lib, "UxTheme.lib")

// ATL definitions and header files
#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

// turns off ATL's hiding of some common and often safely ignored warning messages
#define _ATL_ALL_WARNINGS

#include <AtlBase.h>
#include <AtlWin.h>
#include <AtlCom.h>
#include <AtlCtl.h>
#include <AtlStr.h>
#include <AtlColl.h>
#include <AtlPath.h>

// Additional headers
//#include "resource.h"

// Enable visual styles (Common Controls v6.0)
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
