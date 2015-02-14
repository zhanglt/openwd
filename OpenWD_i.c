

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Wed Feb 25 20:09:17 2015
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


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IComponentRegistrar,0xa817e7a2,0x43fa,0x11d0,0x9e,0x44,0x00,0xaa,0x00,0xb6,0x77,0x0a);


MIDL_DEFINE_GUID(IID, IID_IOpenEdit,0x7C077787,0xC729,0x48E2,0xBB,0x4A,0x00,0xBF,0x00,0xBD,0x27,0x4F);


MIDL_DEFINE_GUID(IID, LIBID_OpenWDLib,0xC0A07342,0xEE8F,0x43A1,0xA6,0xC7,0x81,0x7D,0x49,0xA6,0x94,0x8F);


MIDL_DEFINE_GUID(CLSID, CLSID_CompReg,0x3C78CC30,0xEE07,0x4A63,0x8C,0xA7,0x36,0xAF,0x50,0x73,0xAD,0x17);


MIDL_DEFINE_GUID(IID, DIID__IOpenEditEvents,0x43836F27,0x31CC,0x4523,0x81,0x90,0x09,0xDF,0xE7,0xD0,0x72,0x9B);


MIDL_DEFINE_GUID(CLSID, CLSID_OpenEdit,0x79DAD3A5,0x311C,0x41C5,0x8F,0x57,0xD0,0x83,0xA2,0x93,0x3D,0x2B);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



