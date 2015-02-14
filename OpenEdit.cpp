// OpenEdit.cpp : COpenEdit 的实现

#include "stdafx.h"
#include "OpenEdit.h"
#include "word/msword.h"
#include "word/word.h"
#include "wps/kingsoftWPS.h"
#include "wps/wps.h"


wdocx::CApplication oWordApp; //ms word 
wpsDoc::CApplication oWpsApp;  // kingsoft wps
COleVariant   vOpt(DISP_E_PARAMNOTFOUND, VT_ERROR);
// COpenEdit

STDMETHODIMP COpenEdit::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* const arr[] = 
	{
		&IID_IOpenEdit
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}





STDMETHODIMP COpenEdit::get_DocumentType(int* pDocType)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	*pDocType = this->nDocumentType;

	return S_OK;
}


STDMETHODIMP COpenEdit::put_DocumentType(int nDocType)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	this->nDocumentType = nDocType;

	return S_OK;
}


STDMETHODIMP COpenEdit::GetDocumentFile(BSTR sHeader, BSTR sUserName, int nState, BOOL bTrace)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	// TODO:  在此添加实现代码

	if (this->nDocumentType == 0){//文档类型（ms office）

		if (oWordApp.CreateDispatch("Word.Application")) {//判断客户端是否安装ms word
			oWordApp.Quit(vOpt, vOpt, vOpt);
			oWordApp.ReleaseDispatch(); //ReleaseDispatch()不能关闭当前启动的winword.exe进程，需要使用wordApp.quit() 来退出进程。
			if (!wdocx::GetDocFileFromServer(W2A(sHeader), W2A(sUserName), nState, 1)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}else{//如果没有安装ms word ，启动wps处理

			if (!wpsDoc::GetWpsFileFromServer(W2A(sHeader), W2A(sUserName), 1)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
	}else{//文档类型（Kingsoft office wps）

		if (oWpsApp.CreateDispatch("Wps.Application")) {//首先判断客户端是否安装金山WPS
			oWpsApp.ReleaseDispatch();
			AfxMessageBox("ddddddddddd");
			if (!wpsDoc::GetWpsFileFromServer(W2A(sHeader), W2A(sUserName), 1)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
		else{//如果没有安装金山WPS，启动ms word 来处理

			if (!wdocx::GetDocFileFromServer(W2A(sHeader), W2A(sUserName), nState, 1)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
	}

	AfxGetApp()->DoWaitCursor(0);
	
	return S_OK;
}


STDMETHODIMP COpenEdit::GetAttachment(BSTR sInfo, BSTR sFile, int idx)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码

	return S_OK;
}


STDMETHODIMP COpenEdit::SendDocumentFile(BSTR sHeader, int index)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码

	return S_OK;
}


STDMETHODIMP COpenEdit::SendAttachment(BSTR sInfo)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码

	return S_OK;
}


STDMETHODIMP COpenEdit::get_ServerIp(int* IP)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	*IP = this->nIP;

	return S_OK;
}


STDMETHODIMP COpenEdit::put_ServerIp(int IP)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	this->nIP = IP;

	// TODO:  在此添加实现代码

	return S_OK;
}
STDMETHODIMP COpenEdit::get_ServerPort(int* pPort)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	*pPort = this->nPort;
	return S_OK;
}


STDMETHODIMP COpenEdit::put_ServerPort(int iPort)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	this->nPort = iPort;

	return S_OK;
}

STDMETHODIMP COpenEdit::get_ServerPath(BSTR* pPath)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	*pPath = this->sPath;

	return S_OK;
}


STDMETHODIMP COpenEdit::put_ServerPath(BSTR sPath)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	this->sPath = sPath;

	return S_OK;
}

