#if !defined(AFX_SHOWMSGDLG_H__D2AF4E8F_FE6F_45B6_9669_D674C3B1CD80__INCLUDED_)
#define AFX_SHOWMSGDLG_H__D2AF4E8F_FE6F_45B6_9669_D674C3B1CD80__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ShowMsgDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// ShowMsgDlg dialog
#include "../../Resource.h"

class ShowMsgDlg : public CDialog
{
	// Construction
public:
	ShowMsgDlg(CWnd* pParent = NULL);   // standard constructor
	CString szFileName;
	int nMark;
	// Dialog Data
	//{{AFX_DATA(ShowMsgDlg)
	enum { IDD = IDD_DOWNLOAD };
	CString	m_FileName;
	CString m_Path;
	//}}AFX_DATA


	// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(ShowMsgDlg)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

	// Implementation
protected:
	BOOL IsTheFileExist(CString szFileName);

	// Generated message map functions
	//{{AFX_MSG(ShowMsgDlg)
	virtual void OnOK();
	virtual void OnCancel();
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Kingsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SHOWMSGDLG_H__D2AF4E8F_FE6F_45B6_9669_D674C3B1CD80__INCLUDED_)
