

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Wed Apr 17 13:55:07 2013
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

#ifndef __ppxsolid_i_h__
#define __ppxsolid_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __Ippxsolid_FWD_DEFINED__
#define __Ippxsolid_FWD_DEFINED__
typedef interface Ippxsolid Ippxsolid;
#endif 	/* __Ippxsolid_FWD_DEFINED__ */


#ifndef __ISwDocument_FWD_DEFINED__
#define __ISwDocument_FWD_DEFINED__
typedef interface ISwDocument ISwDocument;
#endif 	/* __ISwDocument_FWD_DEFINED__ */


#ifndef __IDocView_FWD_DEFINED__
#define __IDocView_FWD_DEFINED__
typedef interface IDocView IDocView;
#endif 	/* __IDocView_FWD_DEFINED__ */


#ifndef __IBitmapHandler_FWD_DEFINED__
#define __IBitmapHandler_FWD_DEFINED__
typedef interface IBitmapHandler IBitmapHandler;
#endif 	/* __IBitmapHandler_FWD_DEFINED__ */


#ifndef __IPMPageHandler_FWD_DEFINED__
#define __IPMPageHandler_FWD_DEFINED__
typedef interface IPMPageHandler IPMPageHandler;
#endif 	/* __IPMPageHandler_FWD_DEFINED__ */


#ifndef __IUserPropertyManagerPage_FWD_DEFINED__
#define __IUserPropertyManagerPage_FWD_DEFINED__
typedef interface IUserPropertyManagerPage IUserPropertyManagerPage;
#endif 	/* __IUserPropertyManagerPage_FWD_DEFINED__ */


#ifndef __ppxsolid_FWD_DEFINED__
#define __ppxsolid_FWD_DEFINED__

#ifdef __cplusplus
typedef class ppxsolid ppxsolid;
#else
typedef struct ppxsolid ppxsolid;
#endif /* __cplusplus */

#endif 	/* __ppxsolid_FWD_DEFINED__ */


#ifndef __SwDocument_FWD_DEFINED__
#define __SwDocument_FWD_DEFINED__

#ifdef __cplusplus
typedef class SwDocument SwDocument;
#else
typedef struct SwDocument SwDocument;
#endif /* __cplusplus */

#endif 	/* __SwDocument_FWD_DEFINED__ */


#ifndef __DocView_FWD_DEFINED__
#define __DocView_FWD_DEFINED__

#ifdef __cplusplus
typedef class DocView DocView;
#else
typedef struct DocView DocView;
#endif /* __cplusplus */

#endif 	/* __DocView_FWD_DEFINED__ */


#ifndef __BitmapHandler_FWD_DEFINED__
#define __BitmapHandler_FWD_DEFINED__

#ifdef __cplusplus
typedef class BitmapHandler BitmapHandler;
#else
typedef struct BitmapHandler BitmapHandler;
#endif /* __cplusplus */

#endif 	/* __BitmapHandler_FWD_DEFINED__ */


#ifndef __PMPageHandler_FWD_DEFINED__
#define __PMPageHandler_FWD_DEFINED__

#ifdef __cplusplus
typedef class PMPageHandler PMPageHandler;
#else
typedef struct PMPageHandler PMPageHandler;
#endif /* __cplusplus */

#endif 	/* __PMPageHandler_FWD_DEFINED__ */


#ifndef __UserPropertyManagerPage_FWD_DEFINED__
#define __UserPropertyManagerPage_FWD_DEFINED__

#ifdef __cplusplus
typedef class UserPropertyManagerPage UserPropertyManagerPage;
#else
typedef struct UserPropertyManagerPage UserPropertyManagerPage;
#endif /* __cplusplus */

#endif 	/* __UserPropertyManagerPage_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __Ippxsolid_INTERFACE_DEFINED__
#define __Ippxsolid_INTERFACE_DEFINED__

/* interface Ippxsolid */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_Ippxsolid;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("5F261938-A5A8-41E8-AB0D-58C12EED34C9")
    Ippxsolid : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ToolbarCallback0( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ToolbarEnable0( 
            /* [retval][out] */ long *status) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowPMP( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE EnablePMP( 
            /* [retval][out] */ long *status) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FlyoutCallback( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FlyoutCallback0( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FlyoutCallback1( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FlyoutEnable0( 
            /* [retval][out] */ long *status) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FlyoutEnableCallback0( 
            /* [retval][out] */ long *status) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IppxsolidVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            Ippxsolid * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            Ippxsolid * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            Ippxsolid * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            Ippxsolid * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            Ippxsolid * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            Ippxsolid * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            Ippxsolid * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ToolbarCallback0 )( 
            Ippxsolid * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ToolbarEnable0 )( 
            Ippxsolid * This,
            /* [retval][out] */ long *status);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ShowPMP )( 
            Ippxsolid * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *EnablePMP )( 
            Ippxsolid * This,
            /* [retval][out] */ long *status);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *FlyoutCallback )( 
            Ippxsolid * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *FlyoutCallback0 )( 
            Ippxsolid * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *FlyoutCallback1 )( 
            Ippxsolid * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *FlyoutEnable0 )( 
            Ippxsolid * This,
            /* [retval][out] */ long *status);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *FlyoutEnableCallback0 )( 
            Ippxsolid * This,
            /* [retval][out] */ long *status);
        
        END_INTERFACE
    } IppxsolidVtbl;

    interface Ippxsolid
    {
        CONST_VTBL struct IppxsolidVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define Ippxsolid_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define Ippxsolid_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define Ippxsolid_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define Ippxsolid_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define Ippxsolid_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define Ippxsolid_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define Ippxsolid_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define Ippxsolid_ToolbarCallback0(This)	\
    ( (This)->lpVtbl -> ToolbarCallback0(This) ) 

#define Ippxsolid_ToolbarEnable0(This,status)	\
    ( (This)->lpVtbl -> ToolbarEnable0(This,status) ) 

#define Ippxsolid_ShowPMP(This)	\
    ( (This)->lpVtbl -> ShowPMP(This) ) 

#define Ippxsolid_EnablePMP(This,status)	\
    ( (This)->lpVtbl -> EnablePMP(This,status) ) 

#define Ippxsolid_FlyoutCallback(This)	\
    ( (This)->lpVtbl -> FlyoutCallback(This) ) 

#define Ippxsolid_FlyoutCallback0(This)	\
    ( (This)->lpVtbl -> FlyoutCallback0(This) ) 

#define Ippxsolid_FlyoutCallback1(This)	\
    ( (This)->lpVtbl -> FlyoutCallback1(This) ) 

#define Ippxsolid_FlyoutEnable0(This,status)	\
    ( (This)->lpVtbl -> FlyoutEnable0(This,status) ) 

#define Ippxsolid_FlyoutEnableCallback0(This,status)	\
    ( (This)->lpVtbl -> FlyoutEnableCallback0(This,status) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __Ippxsolid_INTERFACE_DEFINED__ */


#ifndef __ISwDocument_INTERFACE_DEFINED__
#define __ISwDocument_INTERFACE_DEFINED__

/* interface ISwDocument */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ISwDocument;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("57E14605-DE69-4CDB-8BFC-C42E0F0F7FE7")
    ISwDocument : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct ISwDocumentVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISwDocument * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISwDocument * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISwDocument * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISwDocument * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISwDocument * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISwDocument * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISwDocument * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } ISwDocumentVtbl;

    interface ISwDocument
    {
        CONST_VTBL struct ISwDocumentVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISwDocument_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISwDocument_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISwDocument_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISwDocument_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISwDocument_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISwDocument_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISwDocument_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ISwDocument_INTERFACE_DEFINED__ */


#ifndef __IDocView_INTERFACE_DEFINED__
#define __IDocView_INTERFACE_DEFINED__

/* interface IDocView */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IDocView;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("C0DD4BAC-014A-4973-A92C-08F734996ED7")
    IDocView : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IDocViewVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IDocView * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IDocView * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IDocView * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IDocView * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IDocView * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IDocView * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IDocView * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IDocViewVtbl;

    interface IDocView
    {
        CONST_VTBL struct IDocViewVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDocView_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IDocView_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IDocView_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IDocView_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IDocView_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IDocView_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IDocView_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IDocView_INTERFACE_DEFINED__ */


#ifndef __IBitmapHandler_INTERFACE_DEFINED__
#define __IBitmapHandler_INTERFACE_DEFINED__

/* interface IBitmapHandler */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IBitmapHandler;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("28F05B18-4BFF-4BD4-BB16-CA458965C93D")
    IBitmapHandler : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CreateBitmapFileFromResource( 
            /* [in] */ DWORD resID,
            /* [retval][out] */ BSTR *retval) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Dispose( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IBitmapHandlerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IBitmapHandler * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IBitmapHandler * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IBitmapHandler * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IBitmapHandler * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IBitmapHandler * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IBitmapHandler * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IBitmapHandler * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CreateBitmapFileFromResource )( 
            IBitmapHandler * This,
            /* [in] */ DWORD resID,
            /* [retval][out] */ BSTR *retval);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Dispose )( 
            IBitmapHandler * This);
        
        END_INTERFACE
    } IBitmapHandlerVtbl;

    interface IBitmapHandler
    {
        CONST_VTBL struct IBitmapHandlerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IBitmapHandler_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IBitmapHandler_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IBitmapHandler_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IBitmapHandler_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IBitmapHandler_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IBitmapHandler_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IBitmapHandler_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IBitmapHandler_CreateBitmapFileFromResource(This,resID,retval)	\
    ( (This)->lpVtbl -> CreateBitmapFileFromResource(This,resID,retval) ) 

#define IBitmapHandler_Dispose(This)	\
    ( (This)->lpVtbl -> Dispose(This) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IBitmapHandler_INTERFACE_DEFINED__ */


#ifndef __IPMPageHandler_INTERFACE_DEFINED__
#define __IPMPageHandler_INTERFACE_DEFINED__

/* interface IPMPageHandler */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IPMPageHandler;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("E6CA143C-75D5-4BC8-860D-D4C8685B2F3A")
    IPMPageHandler : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IPMPageHandlerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IPMPageHandler * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IPMPageHandler * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IPMPageHandler * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IPMPageHandler * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IPMPageHandler * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IPMPageHandler * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IPMPageHandler * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IPMPageHandlerVtbl;

    interface IPMPageHandler
    {
        CONST_VTBL struct IPMPageHandlerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IPMPageHandler_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IPMPageHandler_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IPMPageHandler_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IPMPageHandler_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IPMPageHandler_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IPMPageHandler_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IPMPageHandler_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IPMPageHandler_INTERFACE_DEFINED__ */


#ifndef __IUserPropertyManagerPage_INTERFACE_DEFINED__
#define __IUserPropertyManagerPage_INTERFACE_DEFINED__

/* interface IUserPropertyManagerPage */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IUserPropertyManagerPage;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("3C8A2C14-8D2B-453C-9A4A-A2CCA398C201")
    IUserPropertyManagerPage : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IUserPropertyManagerPageVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IUserPropertyManagerPage * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IUserPropertyManagerPage * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IUserPropertyManagerPage * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IUserPropertyManagerPage * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IUserPropertyManagerPage * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IUserPropertyManagerPage * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IUserPropertyManagerPage * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IUserPropertyManagerPageVtbl;

    interface IUserPropertyManagerPage
    {
        CONST_VTBL struct IUserPropertyManagerPageVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IUserPropertyManagerPage_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IUserPropertyManagerPage_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IUserPropertyManagerPage_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IUserPropertyManagerPage_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IUserPropertyManagerPage_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IUserPropertyManagerPage_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IUserPropertyManagerPage_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IUserPropertyManagerPage_INTERFACE_DEFINED__ */



#ifndef __ppxsolidLib_LIBRARY_DEFINED__
#define __ppxsolidLib_LIBRARY_DEFINED__

/* library ppxsolidLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ppxsolidLib;

EXTERN_C const CLSID CLSID_ppxsolid;

#ifdef __cplusplus

class DECLSPEC_UUID("CDBC7D74-324B-4BB8-A43C-91F2507E44F1")
ppxsolid;
#endif

EXTERN_C const CLSID CLSID_SwDocument;

#ifdef __cplusplus

class DECLSPEC_UUID("1EB26297-B846-403F-99DF-ED04E19F2865")
SwDocument;
#endif

EXTERN_C const CLSID CLSID_DocView;

#ifdef __cplusplus

class DECLSPEC_UUID("EE2555DC-8D06-4AFE-9F92-C45EB0737971")
DocView;
#endif

EXTERN_C const CLSID CLSID_BitmapHandler;

#ifdef __cplusplus

class DECLSPEC_UUID("992A87C0-31CD-44FD-A5E1-F726995F03D3")
BitmapHandler;
#endif

EXTERN_C const CLSID CLSID_PMPageHandler;

#ifdef __cplusplus

class DECLSPEC_UUID("21099C1D-799F-4E7C-BB5F-BF7FC4C38868")
PMPageHandler;
#endif

EXTERN_C const CLSID CLSID_UserPropertyManagerPage;

#ifdef __cplusplus

class DECLSPEC_UUID("86E1BFBA-6C82-4EE4-A0D5-73AB78AF81F0")
UserPropertyManagerPage;
#endif
#endif /* __ppxsolidLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


