// BrowseDirDialog.h: interface for the CBrowseDirDialog class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BROWSEDIRDIALOG_H__13E6A8A4_C64C_44F9_94E4_6184C9746318__INCLUDED_)
#define AFX_BROWSEDIRDIALOG_H__13E6A8A4_C64C_44F9_94E4_6184C9746318__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBrowseDirDialog
{
public:
	CBrowseDirDialog();
	virtual ~CBrowseDirDialog();
	int DoBrowse();
	CString m_Path;
	CString m_InitDir;
	CString m_SelDir;
	CString m_Title;
	int m_ImageIndex;
};

#endif // !defined(AFX_BROWSEDIRDIALOG_H__13E6A8A4_C64C_44F9_94E4_6184C9746318__INCLUDED_)
