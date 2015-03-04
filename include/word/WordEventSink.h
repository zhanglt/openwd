#if !defined(AFX_WORDEVENTSINK_H__F45A8330_C1E9_11D2_A0C4_0080C7F3B56B__INCLUDED_)
#define AFX_WORDEVENTSINK_H__F45A8330_C1E9_11D2_A0C4_0080C7F3B56B__INCLUDED_

/*----------------------------------------------------------------------------*/

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/*----------------------------------------------------------------------------*/

#pragma warning (disable:4146)
#import "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE15\\MSO.dll"
#pragma warning (default:4146)
//#import "vbeext1.olb"
#import "C:\\Program Files (x86)\\Common Files\\microsoft shared\\VBA\\VBA6\\VBE6EXT.OLB"
#import "C:\\Program Files\\Microsoft Office\\Office15\\MSWORD.OLB" rename("ExitWindows", "WordExitWindows")

#include "ConnectionAdvisor.h"
class  COpenEdit;

/*----------------------------------------------------------------------------*/

const IID IID_IWordAppEventSink = __uuidof(Word::ApplicationEvents);
const IID IID_IWordDocEventSink = __uuidof(Word::DocumentEvents);

/*----------------------------------------------------------------------------*/

class CWordEventSink : public CCmdTarget
{
	DECLARE_DYNCREATE(CWordEventSink)

public:
	CWordEventSink();
	virtual ~CWordEventSink();
	BOOL Advise(IUnknown* pSource, REFIID iid);
	BOOL Unadvise(REFIID iid);
	void SetLauncher(COpenEdit* pWordLauncher);

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWordEventSink)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

protected:

	// Generated message map functions
	//{{AFX_MSG(CWordEventSink)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CWordEventSink)
	afx_msg void OnAppStartup();
	afx_msg void OnAppQuit();
	afx_msg void OnAppDocumentChange();
	afx_msg void OnDocNew();
	afx_msg void OnDocOpen();
	afx_msg void OnDocClose();
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()

private:
	CConnectionAdvisor m_AppEventsAdvisor;
	CConnectionAdvisor m_DocEventsAdvisor;
	COpenEdit* m_pWordLauncher;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WORDEVENTSINK_H__F45A8330_C1E9_11D2_A0C4_0080C7F3B56B__INCLUDED_)
