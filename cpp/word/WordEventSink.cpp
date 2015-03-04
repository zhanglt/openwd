#include "stdafx.h"
#include "OpenEdit.h"
#include "WordEventSink.h"
#include "Regedit.h"

/*----------------------------------------------------------------------------*/

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/*----------------------------------------------------------------------------*/

BEGIN_MESSAGE_MAP(CWordEventSink, CCmdTarget)
	//{{AFX_MSG_MAP(CWordEventSink)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/*----------------------------------------------------------------------------*/

BEGIN_DISPATCH_MAP(CWordEventSink, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CWordEventSink)
	DISP_FUNCTION(CWordEventSink, "Startup",		OnAppStartup,			VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CWordEventSink, "Quit",			OnAppQuit,				VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CWordEventSink, "DocumentChange",	OnAppDocumentChange,	VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CWordEventSink, "New",			OnDocNew,				VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CWordEventSink, "Open",			OnDocOpen,				VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CWordEventSink, "Close",			OnDocClose,				VT_EMPTY, VTS_NONE)
	//DISP_FUNCTION_ID(CWordEventSink, "Startup",			0x01, OnAppStartup,			VT_EMPTY, VTS_NONE)
	//DISP_FUNCTION_ID(CWordEventSink, "Quit",			0x02, OnAppQuit,			VT_EMPTY, VTS_NONE)
	//DISP_FUNCTION_ID(CWordEventSink, "DocumentChange",	0x03, OnAppDocumentChange,	VT_EMPTY, VTS_NONE)
	//DISP_FUNCTION_ID(CWordEventSink, "New",				0x04, OnDocNew,				VT_EMPTY, VTS_NONE)
	//DISP_FUNCTION_ID(CWordEventSink, "Open",			0x05, OnDocOpen,			VT_EMPTY, VTS_NONE)
	//DISP_FUNCTION_ID(CWordEventSink, "Close",			0x06, OnDocClose,			VT_EMPTY, VTS_NONE)

  //}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

/*----------------------------------------------------------------------------*/

BEGIN_INTERFACE_MAP(CWordEventSink, CCmdTarget)
	INTERFACE_PART(CWordEventSink, IID_IWordAppEventSink, Dispatch)
	INTERFACE_PART(CWordEventSink, IID_IWordDocEventSink, Dispatch)
END_INTERFACE_MAP()

/*----------------------------------------------------------------------------*/

IMPLEMENT_DYNCREATE(CWordEventSink, CCmdTarget)

/*----------------------------------------------------------------------------*/

CWordEventSink::CWordEventSink() :
					m_AppEventsAdvisor(IID_IWordAppEventSink), 
					m_DocEventsAdvisor(IID_IWordDocEventSink)
{
	m_pWordLauncher = NULL;
	EnableAutomation();
}

/*----------------------------------------------------------------------------*/

CWordEventSink::~CWordEventSink()
{
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::OnAppStartup() 
{
	// You will never receive this event 
	AfxMessageBox("Quit event received");
	WriteLog("OnAppStartup\n");
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::OnAppQuit() 
{
	//AfxMessageBox("AppQuit event received");
	WriteLog("OnAppQuit\n");
	AfxGetMainWnd()->PostMessage(WM_COMMAND, IDCANCEL, 0L);
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::OnAppDocumentChange() 
{
	//AfxMessageBox("AppDocumentChange event received");
	WriteLog("OnAppDocumentChange\n");
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::OnDocNew() 
{
	//AfxMessageBox("DocNew event received");
	WriteLog("OnDocNew\n");
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::OnDocOpen() 
{
	//AfxMessageBox("DocOpen event received");
	WriteLog("OnDocOpen\n");
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::OnDocClose() 
{
	m_pWordLauncher->OnDocClose();

	
}

/*----------------------------------------------------------------------------*/

BOOL CWordEventSink::Advise(IUnknown* pSource, REFIID iid)
{
	// This GetInterface does not AddRef
	IUnknown* pUnknownSink = GetInterface(&IID_IUnknown);
	if (pUnknownSink == NULL)
	{
		return FALSE;
	}

	if (iid == IID_IWordAppEventSink)
	{
		return m_AppEventsAdvisor.Advise(pUnknownSink, pSource);
	}
	else if (iid == IID_IWordDocEventSink)
	{
		return m_DocEventsAdvisor.Advise(pUnknownSink, pSource);
	}
	else 
	{
		return FALSE;
	}
}

/*----------------------------------------------------------------------------*/
	
BOOL CWordEventSink::Unadvise(REFIID iid)
{
	if (iid == IID_IWordAppEventSink)
	{
		return m_AppEventsAdvisor.Unadvise();
	}
	else if (iid == IID_IWordDocEventSink)
	{
		return m_DocEventsAdvisor.Unadvise();
	}
	else 
	{
		return FALSE;
	}
}

/*----------------------------------------------------------------------------*/

void CWordEventSink::SetLauncher(COpenEdit* pWordLauncher)
{
	m_pWordLauncher = pWordLauncher;
}
