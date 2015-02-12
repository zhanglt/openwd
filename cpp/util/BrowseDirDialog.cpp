// BrowseDirDialog.cpp: implementation of the CBrowseDirDialog class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "../../Include/util/BrowseDirDialog.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBrowseDirDialog::CBrowseDirDialog()
{

}

CBrowseDirDialog::~CBrowseDirDialog()
{

}



static int __stdcall BrowseCtrlCallback(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{
	CBrowseDirDialog* pBrowseDirDialogObj = (CBrowseDirDialog*)lpData;
	if (uMsg == BFFM_INITIALIZED
		&& !pBrowseDirDialogObj->m_SelDir.IsEmpty())
	{
		::SendMessage(hwnd, BFFM_SETSELECTION, TRUE, (LPARAM)(LPCTSTR)(pBrowseDirDialogObj->m_SelDir));
	}
	else // uMsg == BFFM_SELCHANGED 
	{
	}
	return 0;
}

int CBrowseDirDialog::DoBrowse()
{
	LPMALLOC pMalloc;
	if (SHGetMalloc(&pMalloc) != NOERROR){
		return 0;
	}
	BROWSEINFO bInfo;
	LPITEMIDLIST pidl;
	ZeroMemory((PVOID)&bInfo, sizeof (BROWSEINFO));
	if (!m_InitDir.IsEmpty()){

		OLECHAR olePath[MAX_PATH];
		ULONG chEaten;
		ULONG dwAttributes;
		HRESULT hr;
		LPSHELLFOLDER pDesktopFolder;

		if (SUCCEEDED(SHGetDesktopFolder(&pDesktopFolder))){

			MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, m_InitDir.GetBuffer(MAX_PATH), -1, olePath, MAX_PATH);

			m_InitDir.ReleaseBuffer(-1);

			hr = pDesktopFolder->ParseDisplayName(NULL, NULL, olePath, &chEaten, &pidl,
				&dwAttributes);

			if (FAILED(hr)){
				pMalloc->Free(pidl);
				pMalloc->Release();
				return 0;
			}
			bInfo.pidlRoot = pidl;
		}
	}

	bInfo.hwndOwner = NULL;
	bInfo.pszDisplayName = m_Path.GetBuffer(MAX_PATH);
	bInfo.lpszTitle = (m_Title.IsEmpty()) ? "打开" : m_Title;
	bInfo.ulFlags = BIF_RETURNFSANCESTORS | BIF_RETURNONLYFSDIRS;
	bInfo.lpfn = BrowseCtrlCallback; //回调函数地址 
	bInfo.lParam = (LPARAM)this;

	if ((pidl = ::SHBrowseForFolder(&bInfo)) == NULL){
		return 0;
	}

	m_Path.ReleaseBuffer();
	m_ImageIndex = bInfo.iImage;

	if (::SHGetPathFromIDList(pidl, m_Path.GetBuffer(MAX_PATH)) == FALSE){
		pMalloc->Free(pidl);
		pMalloc->Release();
		return 0;
	}
	m_Path.ReleaseBuffer();
	pMalloc->Free(pidl);
	pMalloc->Release();
	return 1;
}
