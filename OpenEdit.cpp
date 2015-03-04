// OpenEdit.cpp : COpenEdit 的实现

#include "stdafx.h"
#include "OpenEdit.h"
#include "word/msword.h"
#include "word/word.h"
#include "wps/kingsoftWPS.h"
#include "wps/wps.h"
#include "word/WordEventSink.h"
#include <iostream>
#include <ostream>
#include <fstream>
#include "comutil.h"
#include <comutil.h>
#include <stdio.h>
#include <comdef.h>

#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "kernel32.lib")



using namespace std;
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


STDMETHODIMP COpenEdit::GetDocumentFile(int nOpenMode, BOOL bTrace)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	// TODO:  在此添加实现代码
	//CString Header = "d1";
	//CString test((LPCTSTR)sHeader);

	if (this->nDocumentType == 0){//文档类型（ms office）
		/*
		if (oWordApp.CreateDispatch("Word.Application")) {//判断客户端是否安装ms word
			oWordApp.Quit(vOpt, vOpt, vOpt);
			oWordApp.ReleaseDispatch(); //ReleaseDispatch()不能关闭当前启动的winword.exe进程，需要使用wordApp.quit() 来退出进程。
			if (!wdocx::GetDocFileFromServer(this,W2A(this->sFileID), W2A(this->sUserName), nOpenMode, bTrace)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
			
		}else{//如果没有安装ms word ，启动wps处理
			if (!wpsDoc::GetWpsFileFromServer(W2A(this->sFileID), W2A(this->sUserName), bTrace)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}*/
		HRESULT hr = m_pWord.CreateInstance(__uuidof(Word::Application));
		if (SUCCEEDED(hr)){
			BOOL Res = m_WordEventSink.Advise(m_pWord, IID_IWordAppEventSink);
			if (!wdocx::GetDocFileFromServer(m_pWord, W2A(this->sFileID), W2A(this->sUserName), nOpenMode, bTrace)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
		else{

	
		
		
		
		}

		//	m_pDoc = m_pWord->Documents->Add();
		
		
		
		//Res = m_WordEventSink.Advise(m_pDoc, IID_IWordDocEventSink);
		
		//m_pWord->Visible = VARIANT_TRUE;



	}else{//文档类型（Kingsoft office wps）
	
		if (oWpsApp.CreateDispatch("KWPS.Application")) {//首先判断客户端是否安装金山WPS
			oWpsApp.Quit(vOpt, vOpt, vOpt);
			oWpsApp.ReleaseDispatch();
			if (!wpsDoc::GetWpsFileFromServer(W2A(this->sFileID), W2A(this->sUserName), bTrace)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
		else{//如果没有安装金山WPS，启动ms word 来处理
		
			if (!wdocx::GetDocFileFromServer(this,W2A(this->sFileID), W2A(this->sUserName), nOpenMode, bTrace)) {
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


STDMETHODIMP COpenEdit::SendDocumentFile(BSTR sHeader, int nOpenMode)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	// TODO:  在此添加实现代码
	AfxGetApp()->DoWaitCursor(1);
	if (this->nDocumentType == 0){//文档类型（ms office）

		if (oWordApp.CreateDispatch("Word.Application")) {//判断客户端是否安装ms word
			oWordApp.Quit(vOpt, vOpt, vOpt);
			oWordApp.ReleaseDispatch(); //ReleaseDispatch()不能关闭当前启动的winword.exe进程，需要使用wordApp.quit() 来退出进程。
			if (!wdocx::SendDocFileToServer(W2A(sHeader), nOpenMode)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
		else{//如果没有安装ms word ，启动wps处理
			if (!wpsDoc::SendWpsFileToServer(W2A(sHeader), nOpenMode)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
	}
	else{//文档类型（Kingsoft office wps）

		if (oWpsApp.CreateDispatch("KWPS.Application")) {//首先判断客户端是否安装金山WPS
			oWpsApp.Quit(vOpt, vOpt, vOpt);
			oWpsApp.ReleaseDispatch();
			if (!wpsDoc::SendWpsFileToServer(W2A(sHeader), nOpenMode)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
		else{//如果没有安装金山WPS，启动ms word 来处理

			if (!wdocx::SendDocFileToServer(W2A(sHeader), nOpenMode)) {
				AfxGetApp()->DoWaitCursor(0);
				return S_FALSE;
			}
		}
	}

	AfxGetApp()->DoWaitCursor(0);

	return S_OK;
}


STDMETHODIMP COpenEdit::SendAttachment(BSTR sInfo)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码

	return S_OK;
}


STDMETHODIMP COpenEdit::get_ServerIp(BSTR* IP)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	//*IP = this->nIP;
	CString str;
	::GetProfileString("openwd", "Ip", "127.0.0.1", str.GetBuffer(15),15);
	str.ReleaseBuffer();
	*IP = str.AllocSysString();
	SysFreeString(*IP);
	return S_OK;
}


STDMETHODIMP COpenEdit::put_ServerIp(BSTR IP)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	this->nIP = IP;
	::WriteProfileStringA("openwd", "Ip", (LPCSTR)_bstr_t(IP));
	
	// TODO:  在此添加实现代码
	return S_OK;
}
STDMETHODIMP COpenEdit::get_ServerPort(BSTR* pPort)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	//*pPort = this->nPort;
	CString str;
	::GetProfileString("openwd", "Port", "80", str.GetBuffer(6), 6);
	str.ReleaseBuffer();
	*pPort = str.AllocSysString();
	SysFreeString(*pPort);
	return S_OK;
}


STDMETHODIMP COpenEdit::put_ServerPort(BSTR iPort)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	this->nPort = iPort;

	::WriteProfileString("openwd", "Port", (LPCSTR)_bstr_t(iPort));
	
	return S_OK;
}

STDMETHODIMP COpenEdit::get_ServerPath(BSTR* pPath)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
		// TODO:  在此添加实现代码
	//*pPath = this->sPath;
	CString str;
	::GetProfileString("openwd", "ServerURL", "jc/legalDoc", str.GetBuffer(50), 50);
	str.ReleaseBuffer();
	*pPath = str.AllocSysString();
	SysFreeString(*pPath);
	return S_OK;
}


STDMETHODIMP COpenEdit::put_ServerPath(BSTR sPath)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	this->sPath = sPath;
	::WriteProfileString("openwd", "ServerURL", (LPCSTR)_bstr_t(sPath));
	return S_OK;
}


STDMETHODIMP COpenEdit::get_FileID(BSTR* pFileID)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	// TODO:  在此添加实现代码
	*pFileID = this->sFileID;
	return S_OK;
}


STDMETHODIMP COpenEdit::put_FileID(BSTR sFileID)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	this->sFileID = sFileID;
	return S_OK;
}

STDMETHODIMP COpenEdit::get_UserName(BSTR* pUserName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	// TODO:  在此添加实现代码
	*pUserName = this->sUserName;
	return S_OK;
}


STDMETHODIMP COpenEdit::put_UserName(BSTR sUserName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO:  在此添加实现代码
	this->sUserName = sUserName;
	return S_OK;
}



STDMETHODIMP COpenEdit::ShowWindows(BSTR sTitle, int nCmdShow)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	// TODO:  在此添加实现代码
	ShowWinEx(W2A(sTitle), nCmdShow);
	return S_OK;
}
void COpenEdit::DumpComError(const _com_error& e) const
{
	CString ComErrorMessage;
	ComErrorMessage.Format("COM Error1111: 0x%08lX. %s", e.Error(), e.ErrorMessage());
	AfxMessageBox(ComErrorMessage);
}
void COpenEdit::DumpOleError(const COleException& e) const
{
	CString OleErrorMessage;
	OleErrorMessage.Format("Ole Error : 0x%08lX", (long)e.m_sc);
	AfxMessageBox(OleErrorMessage);
}void COpenEdit::DumpDispatchError(const COleDispatchException& e) const
{
	AfxMessageBox("Dispatch Error : " + e.m_strDescription);
}
BOOL COpenEdit::GetPageCount(DWORD& PageCount)
{
	// To get the page count, you must first get the BuiltInDocumentProperties
	// IDispatch interface, so you have access to the VB collection of document
	// properties. What you have to do next is to get Item(n).Value where n is 
	// the index of the property you want to retrieve (wdPropertyPage here).
	// Don't forget to trap COleException, COleDispatchException and _com_error

	try
	{
		IDispatchPtr pDispatch(m_pWord->ActiveDocument->BuiltInDocumentProperties);
		AfxMessageBox("000000000000000000");
		ASSERT(pDispatch != NULL);
		AfxMessageBox("111111111111111111");
		// this pDispatch will be released by the smart pointer, so use FALSE  
		COleDispatchDriver DocProperties(pDispatch, FALSE);
		_variant_t Property((long)Word::wdPropertyPages);
		_variant_t Result;

		// The Item method is the default member for the collection object
		DocProperties.InvokeHelper(DISPID_VALUE,
			DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			VT_VARIANT,
			(void*)&Result,
			(BYTE*)VTS_VARIANT,
			&Property);
		AfxMessageBox("22222222222222222222222");
		// pDispatch will be extracted from variant Result
		COleDispatchDriver DocProperty(Result);
		// The Value property is the default member for the Item object
		DocProperty.GetProperty(DISPID_VALUE, VT_I4, &PageCount);

	}
	catch (_com_error& ComError)
	{
		DumpComError(ComError);
		return FALSE;
	}

	catch (COleException* pOleError)
	{
		DumpOleError(*pOleError);
		pOleError->Delete();
		return FALSE;
	}

	catch (COleDispatchException* pDispatchError)
	{
		DumpDispatchError(*pDispatchError);
		pDispatchError->Delete();
		return FALSE;
	}

	catch (...)
	{
		return FALSE;
	}

	return TRUE;
}

void COpenEdit::OnDocClose()
{
	DWORD PageCount;
	GetPageCount(PageCount);
	CString Msg;
	Msg.Format("Close event received\nNumber of pages is %d", PageCount);
	AfxMessageBox(Msg);

}