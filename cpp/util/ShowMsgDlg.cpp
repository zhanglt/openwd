// ShowMsgDlg.cpp : implementation file
//

#include "stdafx.h"
#include  <fstream>
#include "util/ShowMsgDlg.h"
using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// ShowMsgDlg dialog


ShowMsgDlg::ShowMsgDlg(CWnd* pParent /*=NULL*/)
: CDialog(ShowMsgDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(ShowMsgDlg)
	m_FileName = _T("");
	//}}AFX_DATA_INIT
	nMark = 0;
}


void ShowMsgDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(ShowMsgDlg)
	DDX_Text(pDX, IDC_EDITNAME, m_FileName);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(ShowMsgDlg, CDialog)
	//{{AFX_MSG_MAP(ShowMsgDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// ShowMsgDlg message handlers
//����
void ShowMsgDlg::OnOK()
{
	// TODO: Add extra validation here
	UpdateData(true);

	if (szFileName != m_FileName){
		if (IsTheFileExist(m_Path + "\\" + m_FileName)){
			MessageBox("���ļ��Ѵ��������������!", "ϵͳ��Ϣ", MB_OK | MB_ICONINFORMATION);
			return;
		}
		nMark = 2;  //�ļ��������ı�
	}else{
		nMark = 1;
	}
	szFileName = m_FileName;  //�����º���ļ�������
	CDialog::OnOK();
}
//����
void ShowMsgDlg::OnCancel()
{
	// TODO: Add extra cleanup here
	nMark = 0;
	CDialog::OnCancel();
}

BOOL ShowMsgDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	m_FileName = szFileName;
	UpdateData(false);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

BOOL ShowMsgDlg::IsTheFileExist(CString szFileName)
{
	ofstream pfile(szFileName, ios_base::out | ios_base::app | ios_base::binary);

	if (!pfile.is_open())  return false;
	pfile.close();
	return true;
}