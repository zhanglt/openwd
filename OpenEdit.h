// OpenEdit.h : COpenEdit 的声明

#pragma once
#include "resource.h"       // 主符号



#include "OpenWD_i.h"
#include "_IOpenEditEvents_CP.h"
#include <string>
#include <iostream>


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;
using namespace std;


// COpenEdit

class ATL_NO_VTABLE COpenEdit :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<COpenEdit, &CLSID_OpenEdit>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<COpenEdit>,
	public CProxy_IOpenEditEvents<COpenEdit>,
	public IObjectWithSiteImpl<COpenEdit>,
	public IDispatchImpl<IOpenEdit, &IID_IOpenEdit, &LIBID_OpenWDLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
private:
	int  nDocumentType;
	int  nIP;
	int  nPort;
	BSTR  sPath;

public:
	COpenEdit()
	{
		//设置默认处理文档类型（默认为=0 为Ms word 类型）
		this->nDocumentType = 0;
		
	}

DECLARE_REGISTRY_RESOURCEID(IDR_OPENEDIT)


BEGIN_COM_MAP(COpenEdit)
	COM_INTERFACE_ENTRY(IOpenEdit)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY(IObjectWithSite)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(COpenEdit)
	CONNECTION_POINT_ENTRY(__uuidof(_IOpenEditEvents))
END_CONNECTION_POINT_MAP()
// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
	


	//************************************
	//word/wps文档标识
	//************************************
	STDMETHOD(get_DocumentType)(int* pVal);
	STDMETHOD(put_DocumentType)(int newVal);
	//************************************
	// Method:    GetDocumentFile
	// FullName:  COpenEdit::GetDocumentFile
	// Access:    public 
	// Returns:   STDMETHODIMP
	// Qualifier:
	// Parameter: BSTR szHeader
	// Parameter: BSTR szUserName
	// Parameter: BOOL bTrace  修改痕迹标识
	//************************************
	STDMETHOD(GetDocumentFile)(BSTR sHeader, BSTR sUserName, BOOL bTrace);

	//************************************
	// Method:    GetAttachment
	// FullName:  COpenEdit::GetAttachment
	// Access:    public 
	// Returns:   STDMETHODIMP
	// Qualifier:
	// Parameter: BSTR sInfo
	// Parameter: BSTR sFile
	// Parameter: BOOL idx 单/多文件标识
	//************************************
	STDMETHOD(GetAttachment)(BSTR sInfo, BSTR sFile, int idx);
	STDMETHOD(SendDocumentFile)(BSTR sHeader, int index);
	STDMETHOD(SendAttachment)(BSTR sInfo);

	STDMETHOD(get_ServerIp)(int* IP);
	STDMETHOD(put_ServerIp)(int IP);
	STDMETHOD(get_ServerPort)(int* pPort);
	STDMETHOD(put_ServerPort)(int iPort);
	STDMETHOD(get_ServerPath)(BSTR* pPath);
	STDMETHOD(put_ServerPath)(BSTR sPath);
};

OBJECT_ENTRY_AUTO(__uuidof(OpenEdit), COpenEdit)
