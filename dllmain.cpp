// dllmain.cpp : DllMain 的实现。

#include "stdafx.h"
#include "resource.h"
#include "OpenWD_i.h"
#include "dllmain.h"
#include "compreg.h"
#include "xdlldata.h"

COpenWDModule _AtlModule;

class COpenWDApp : public CWinApp
{
public:

// 重写
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(COpenWDApp, CWinApp)
END_MESSAGE_MAP()

COpenWDApp theApp;

BOOL COpenWDApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(m_hInstance, DLL_PROCESS_ATTACH, NULL))
		return FALSE;
#endif
	return CWinApp::InitInstance();
}

int COpenWDApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
