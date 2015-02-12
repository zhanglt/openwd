

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Thu Feb 12 22:56:51 2015
 */
/* Compiler settings for OpenWD.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __OpenWD_i_h__
#define __OpenWD_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IComponentRegistrar_FWD_DEFINED__
#define __IComponentRegistrar_FWD_DEFINED__
typedef interface IComponentRegistrar IComponentRegistrar;

#endif 	/* __IComponentRegistrar_FWD_DEFINED__ */


#ifndef __IOpenEdit_FWD_DEFINED__
#define __IOpenEdit_FWD_DEFINED__
typedef interface IOpenEdit IOpenEdit;

#endif 	/* __IOpenEdit_FWD_DEFINED__ */


#ifndef __CompReg_FWD_DEFINED__
#define __CompReg_FWD_DEFINED__

#ifdef __cplusplus
typedef class CompReg CompReg;
#else
typedef struct CompReg CompReg;
#endif /* __cplusplus */

#endif 	/* __CompReg_FWD_DEFINED__ */


#ifndef ___IOpenEditEvents_FWD_DEFINED__
#define ___IOpenEditEvents_FWD_DEFINED__
typedef interface _IOpenEditEvents _IOpenEditEvents;

#endif 	/* ___IOpenEditEvents_FWD_DEFINED__ */


#ifndef __OpenEdit_FWD_DEFINED__
#define __OpenEdit_FWD_DEFINED__

#ifdef __cplusplus
typedef class OpenEdit OpenEdit;
#else
typedef struct OpenEdit OpenEdit;
#endif /* __cplusplus */

#endif 	/* __OpenEdit_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IComponentRegistrar_INTERFACE_DEFINED__
#define __IComponentRegistrar_INTERFACE_DEFINED__

/* interface IComponentRegistrar */
/* [unique][dual][uuid][object] */ 


EXTERN_C const IID IID_IComponentRegistrar;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("a817e7a2-43fa-11d0-9e44-00aa00b6770a")
    IComponentRegistrar : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Attach( 
            /* [in] */ BSTR bstrPath) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE RegisterAll( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UnregisterAll( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetComponents( 
            /* [out] */ SAFEARRAY * *pbstrCLSIDs,
            /* [out] */ SAFEARRAY * *pbstrDescriptions) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE RegisterComponent( 
            /* [in] */ BSTR bstrCLSID) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UnregisterComponent( 
            /* [in] */ BSTR bstrCLSID) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IComponentRegistrarVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IComponentRegistrar * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IComponentRegistrar * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IComponentRegistrar * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IComponentRegistrar * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IComponentRegistrar * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IComponentRegistrar * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IComponentRegistrar * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Attach )( 
            IComponentRegistrar * This,
            /* [in] */ BSTR bstrPath);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *RegisterAll )( 
            IComponentRegistrar * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *UnregisterAll )( 
            IComponentRegistrar * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetComponents )( 
            IComponentRegistrar * This,
            /* [out] */ SAFEARRAY * *pbstrCLSIDs,
            /* [out] */ SAFEARRAY * *pbstrDescriptions);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *RegisterComponent )( 
            IComponentRegistrar * This,
            /* [in] */ BSTR bstrCLSID);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *UnregisterComponent )( 
            IComponentRegistrar * This,
            /* [in] */ BSTR bstrCLSID);
        
        END_INTERFACE
    } IComponentRegistrarVtbl;

    interface IComponentRegistrar
    {
        CONST_VTBL struct IComponentRegistrarVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IComponentRegistrar_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IComponentRegistrar_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IComponentRegistrar_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IComponentRegistrar_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IComponentRegistrar_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IComponentRegistrar_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IComponentRegistrar_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IComponentRegistrar_Attach(This,bstrPath)	\
    ( (This)->lpVtbl -> Attach(This,bstrPath) ) 

#define IComponentRegistrar_RegisterAll(This)	\
    ( (This)->lpVtbl -> RegisterAll(This) ) 

#define IComponentRegistrar_UnregisterAll(This)	\
    ( (This)->lpVtbl -> UnregisterAll(This) ) 

#define IComponentRegistrar_GetComponents(This,pbstrCLSIDs,pbstrDescriptions)	\
    ( (This)->lpVtbl -> GetComponents(This,pbstrCLSIDs,pbstrDescriptions) ) 

#define IComponentRegistrar_RegisterComponent(This,bstrCLSID)	\
    ( (This)->lpVtbl -> RegisterComponent(This,bstrCLSID) ) 

#define IComponentRegistrar_UnregisterComponent(This,bstrCLSID)	\
    ( (This)->lpVtbl -> UnregisterComponent(This,bstrCLSID) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IComponentRegistrar_INTERFACE_DEFINED__ */


#ifndef __IOpenEdit_INTERFACE_DEFINED__
#define __IOpenEdit_INTERFACE_DEFINED__

/* interface IOpenEdit */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IOpenEdit;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("7C077787-C729-48E2-BB4A-00BF00BD274F")
    IOpenEdit : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_DocumentType( 
            /* [retval][out] */ int *pDocType) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_DocumentType( 
            /* [in] */ int nDocType) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetDocumentFile( 
            /* [in] */ BSTR sHeader,
            BSTR sUserName,
            /* [in] */ BOOL bTrace) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetAttachment( 
            /* [in] */ BSTR sInfo,
            /* [in] */ BSTR sFile,
            /* [in] */ int idx) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE PutDocumentFile( 
            /* [in] */ BSTR sHeader,
            /* [in] */ int index) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SendAttachment( 
            /* [in] */ BSTR sInfo) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ServerIp( 
            /* [retval][out] */ int *IP) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ServerIp( 
            /* [in] */ int IP) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ServerPort( 
            /* [retval][out] */ int *pPort) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ServerPort( 
            /* [in] */ int iPort) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ServerPath( 
            /* [retval][out] */ BSTR *pPath) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ServerPath( 
            /* [in] */ BSTR sPath) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IOpenEditVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IOpenEdit * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IOpenEdit * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IOpenEdit * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IOpenEdit * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IOpenEdit * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IOpenEdit * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IOpenEdit * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_DocumentType )( 
            IOpenEdit * This,
            /* [retval][out] */ int *pDocType);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_DocumentType )( 
            IOpenEdit * This,
            /* [in] */ int nDocType);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetDocumentFile )( 
            IOpenEdit * This,
            /* [in] */ BSTR sHeader,
            BSTR sUserName,
            /* [in] */ BOOL bTrace);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetAttachment )( 
            IOpenEdit * This,
            /* [in] */ BSTR sInfo,
            /* [in] */ BSTR sFile,
            /* [in] */ int idx);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *PutDocumentFile )( 
            IOpenEdit * This,
            /* [in] */ BSTR sHeader,
            /* [in] */ int index);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *SendAttachment )( 
            IOpenEdit * This,
            /* [in] */ BSTR sInfo);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ServerIp )( 
            IOpenEdit * This,
            /* [retval][out] */ int *IP);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ServerIp )( 
            IOpenEdit * This,
            /* [in] */ int IP);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ServerPort )( 
            IOpenEdit * This,
            /* [retval][out] */ int *pPort);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ServerPort )( 
            IOpenEdit * This,
            /* [in] */ int iPort);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ServerPath )( 
            IOpenEdit * This,
            /* [retval][out] */ BSTR *pPath);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ServerPath )( 
            IOpenEdit * This,
            /* [in] */ BSTR sPath);
        
        END_INTERFACE
    } IOpenEditVtbl;

    interface IOpenEdit
    {
        CONST_VTBL struct IOpenEditVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOpenEdit_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IOpenEdit_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IOpenEdit_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IOpenEdit_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IOpenEdit_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IOpenEdit_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IOpenEdit_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IOpenEdit_get_DocumentType(This,pDocType)	\
    ( (This)->lpVtbl -> get_DocumentType(This,pDocType) ) 

#define IOpenEdit_put_DocumentType(This,nDocType)	\
    ( (This)->lpVtbl -> put_DocumentType(This,nDocType) ) 

#define IOpenEdit_GetDocumentFile(This,sHeader,sUserName,bTrace)	\
    ( (This)->lpVtbl -> GetDocumentFile(This,sHeader,sUserName,bTrace) ) 

#define IOpenEdit_GetAttachment(This,sInfo,sFile,idx)	\
    ( (This)->lpVtbl -> GetAttachment(This,sInfo,sFile,idx) ) 

#define IOpenEdit_PutDocumentFile(This,sHeader,index)	\
    ( (This)->lpVtbl -> PutDocumentFile(This,sHeader,index) ) 

#define IOpenEdit_SendAttachment(This,sInfo)	\
    ( (This)->lpVtbl -> SendAttachment(This,sInfo) ) 

#define IOpenEdit_get_ServerIp(This,IP)	\
    ( (This)->lpVtbl -> get_ServerIp(This,IP) ) 

#define IOpenEdit_put_ServerIp(This,IP)	\
    ( (This)->lpVtbl -> put_ServerIp(This,IP) ) 

#define IOpenEdit_get_ServerPort(This,pPort)	\
    ( (This)->lpVtbl -> get_ServerPort(This,pPort) ) 

#define IOpenEdit_put_ServerPort(This,iPort)	\
    ( (This)->lpVtbl -> put_ServerPort(This,iPort) ) 

#define IOpenEdit_get_ServerPath(This,pPath)	\
    ( (This)->lpVtbl -> get_ServerPath(This,pPath) ) 

#define IOpenEdit_put_ServerPath(This,sPath)	\
    ( (This)->lpVtbl -> put_ServerPath(This,sPath) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IOpenEdit_INTERFACE_DEFINED__ */



#ifndef __OpenWDLib_LIBRARY_DEFINED__
#define __OpenWDLib_LIBRARY_DEFINED__

/* library OpenWDLib */
/* [custom][version][uuid] */ 


EXTERN_C const IID LIBID_OpenWDLib;

EXTERN_C const CLSID CLSID_CompReg;

#ifdef __cplusplus

class DECLSPEC_UUID("3C78CC30-EE07-4A63-8CA7-36AF5073AD17")
CompReg;
#endif

#ifndef ___IOpenEditEvents_DISPINTERFACE_DEFINED__
#define ___IOpenEditEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IOpenEditEvents */
/* [uuid] */ 


EXTERN_C const IID DIID__IOpenEditEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("43836F27-31CC-4523-8190-09DFE7D0729B")
    _IOpenEditEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IOpenEditEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IOpenEditEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IOpenEditEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IOpenEditEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IOpenEditEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IOpenEditEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IOpenEditEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IOpenEditEvents * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _IOpenEditEventsVtbl;

    interface _IOpenEditEvents
    {
        CONST_VTBL struct _IOpenEditEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IOpenEditEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IOpenEditEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IOpenEditEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IOpenEditEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IOpenEditEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IOpenEditEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IOpenEditEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IOpenEditEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_OpenEdit;

#ifdef __cplusplus

class DECLSPEC_UUID("79DAD3A5-311C-41C5-8F57-D083A2933D2B")
OpenEdit;
#endif
#endif /* __OpenWDLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree(     unsigned long *, LPSAFEARRAY * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


