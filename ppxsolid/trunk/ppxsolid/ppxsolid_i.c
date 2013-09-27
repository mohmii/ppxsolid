

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Tue Sep 24 10:53:04 2013
 */
/* Compiler settings for ppxsolid.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
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

MIDL_DEFINE_GUID(IID, IID_Ippxsolid,0x5F261938,0xA5A8,0x41E8,0xAB,0x0D,0x58,0xC1,0x2E,0xED,0x34,0xC9);


MIDL_DEFINE_GUID(IID, IID_ISwDocument,0x57E14605,0xDE69,0x4CDB,0x8B,0xFC,0xC4,0x2E,0x0F,0x0F,0x7F,0xE7);


MIDL_DEFINE_GUID(IID, IID_IDocView,0xC0DD4BAC,0x014A,0x4973,0xA9,0x2C,0x08,0xF7,0x34,0x99,0x6E,0xD7);


MIDL_DEFINE_GUID(IID, IID_IBitmapHandler,0x28F05B18,0x4BFF,0x4BD4,0xBB,0x16,0xCA,0x45,0x89,0x65,0xC9,0x3D);


MIDL_DEFINE_GUID(IID, IID_IPMPageHandler,0xE6CA143C,0x75D5,0x4BC8,0x86,0x0D,0xD4,0xC8,0x68,0x5B,0x2F,0x3A);


MIDL_DEFINE_GUID(IID, IID_IUserPropertyManagerPage,0x3C8A2C14,0x8D2B,0x453C,0x9A,0x4A,0xA2,0xCC,0xA3,0x98,0xC2,0x01);


MIDL_DEFINE_GUID(IID, LIBID_ppxsolidLib,0xF55838FE,0xF5A2,0x44FE,0xB0,0x23,0x8E,0x2A,0xD7,0x82,0xB2,0x97);


MIDL_DEFINE_GUID(CLSID, CLSID_ppxsolid,0xCDBC7D74,0x324B,0x4BB8,0xA4,0x3C,0x91,0xF2,0x50,0x7E,0x44,0xF1);


MIDL_DEFINE_GUID(CLSID, CLSID_SwDocument,0x1EB26297,0xB846,0x403F,0x99,0xDF,0xED,0x04,0xE1,0x9F,0x28,0x65);


MIDL_DEFINE_GUID(CLSID, CLSID_DocView,0xEE2555DC,0x8D06,0x4AFE,0x9F,0x92,0xC4,0x5E,0xB0,0x73,0x79,0x71);


MIDL_DEFINE_GUID(CLSID, CLSID_BitmapHandler,0x992A87C0,0x31CD,0x44FD,0xA5,0xE1,0xF7,0x26,0x99,0x5F,0x03,0xD3);


MIDL_DEFINE_GUID(CLSID, CLSID_PMPageHandler,0x21099C1D,0x799F,0x4E7C,0xBB,0x5F,0xBF,0x7F,0xC4,0xC3,0x88,0x68);


MIDL_DEFINE_GUID(CLSID, CLSID_UserPropertyManagerPage,0x86E1BFBA,0x6C82,0x4EE4,0xA0,0xD5,0x73,0xAB,0x78,0xAF,0x81,0xF0);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



