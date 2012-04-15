// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "TRContactFinder_i.h"
#include "dllmain.h"

CTRContactFinderModule _AtlModule;

class CTRContactFinderApp : public CWinApp
{
public:

// Overrides
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CTRContactFinderApp, CWinApp)
END_MESSAGE_MAP()

CTRContactFinderApp theApp;

BOOL CTRContactFinderApp::InitInstance()
{
	return CWinApp::InitInstance();
}

int CTRContactFinderApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
