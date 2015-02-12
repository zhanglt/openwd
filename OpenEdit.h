// OpenEdit.h : COpenEdit ������

#pragma once
#include "resource.h"       // ������



#include "OpenWD_i.h"
#include "_IOpenEditEvents_CP.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE ƽ̨(�粻�ṩ��ȫ DCOM ֧�ֵ� Windows Mobile ƽ̨)���޷���ȷ֧�ֵ��߳� COM ���󡣶��� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ��ǿ�� ATL ֧�ִ������߳� COM ����ʵ�ֲ�����ʹ���䵥�߳� COM ����ʵ�֡�rgs �ļ��е��߳�ģ���ѱ�����Ϊ��Free����ԭ���Ǹ�ģ���Ƿ� DCOM Windows CE ƽ̨֧�ֵ�Ψһ�߳�ģ�͡�"
#endif

using namespace ATL;


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

public:
	COpenEdit()
	{
		//����Ĭ�ϴ����ĵ����ͣ�Ĭ��Ϊ=0 ΪMs word ���ͣ�
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
	//word/wps�ĵ���ʶ
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
	// Parameter: BOOL bTrace  �޸ĺۼ���ʶ
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
	// Parameter: BOOL idx ��/���ļ���ʶ
	//************************************
	STDMETHOD(GetAttachment)(BSTR sInfo, BSTR sFile, int idx);
	STDMETHOD(PutDocumentFile)(BSTR sHeader, int index);
};

OBJECT_ENTRY_AUTO(__uuidof(OpenEdit), COpenEdit)