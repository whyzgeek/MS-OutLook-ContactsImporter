// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif
//
//#ifndef _AFXDLL
//#define _AFXDLL
//#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit
#define VC_EXTRALEAN	
#include <afxwin.h>
#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
//#define SITELOCK_USE_IOLEOBJECT
//#define SITELOCK_SUPPORT_DOWNLOAD
//#include "sitelock.h"

using namespace ATL;

class CTRError: public CObject
{
public:
	SHORT code;
	CString err;
};
