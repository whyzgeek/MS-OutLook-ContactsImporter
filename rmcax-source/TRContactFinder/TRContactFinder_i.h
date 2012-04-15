

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Sat Sep 19 19:38:28 2009
 */
/* Compiler settings for .\TRContactFinder.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

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

#ifndef __TRContactFinder_i_h__
#define __TRContactFinder_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IContactFinderCtrl_FWD_DEFINED__
#define __IContactFinderCtrl_FWD_DEFINED__
typedef interface IContactFinderCtrl IContactFinderCtrl;
#endif 	/* __IContactFinderCtrl_FWD_DEFINED__ */


#ifndef ___IContactFinderCtrlEvents_FWD_DEFINED__
#define ___IContactFinderCtrlEvents_FWD_DEFINED__
typedef interface _IContactFinderCtrlEvents _IContactFinderCtrlEvents;
#endif 	/* ___IContactFinderCtrlEvents_FWD_DEFINED__ */


#ifndef __ContactFinderCtrl_FWD_DEFINED__
#define __ContactFinderCtrl_FWD_DEFINED__

#ifdef __cplusplus
typedef class ContactFinderCtrl ContactFinderCtrl;
#else
typedef struct ContactFinderCtrl ContactFinderCtrl;
#endif /* __cplusplus */

#endif 	/* __ContactFinderCtrl_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IContactFinderCtrl_INTERFACE_DEFINED__
#define __IContactFinderCtrl_INTERFACE_DEFINED__

/* interface IContactFinderCtrl */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IContactFinderCtrl;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("CE2074C6-86AB-4f46-95F8-9B4B2088E373")
    IContactFinderCtrl : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE InitCtrl( 
            /* [retval][out] */ ULONG *pInit) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetPABContactsCount( 
            /* [retval][out] */ ULONG *pCount) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetNK2ContactsCount( 
            /* [retval][out] */ ULONG *pCount) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetPABFolderCount( 
            /* [retval][out] */ ULONG *pCount) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StartPABContactSearch( 
            /* [in] */ BSTR bstrSearchBy,
            /* [in] */ BSTR bstrStartWith,
            /* [in] */ ULONG lPageNumber) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StartNK2ContactSearch( 
            /* [in] */ BSTR bstrSearchBy,
            /* [in] */ BSTR bstrStartWith,
            /* [in] */ ULONG lPageNumber) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Abort( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AbortPABSearch( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AbortNK2Search( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StartPABFolderCounting( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StartPABContactCounting( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AbortPABFolderCounting( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AbortPABContactCounting( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_FetchUnique( 
            /* [retval][out] */ USHORT *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_FetchUnique( 
            /* [in] */ USHORT newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PABPageSize( 
            /* [retval][out] */ ULONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_PABPageSize( 
            /* [in] */ ULONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_NK2PageSize( 
            /* [retval][out] */ ULONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_NK2PageSize( 
            /* [in] */ ULONG newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StartNK2ContactCounting( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AbortNK2ContactCounting( void) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_SubFolderLevel( 
            /* [in] */ ULONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableLog( 
            /* [retval][out] */ USHORT *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableLog( 
            /* [in] */ USHORT newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IContactFinderCtrlVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IContactFinderCtrl * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IContactFinderCtrl * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IContactFinderCtrl * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IContactFinderCtrl * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IContactFinderCtrl * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IContactFinderCtrl * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IContactFinderCtrl * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *InitCtrl )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ ULONG *pInit);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetPABContactsCount )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ ULONG *pCount);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetNK2ContactsCount )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ ULONG *pCount);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetPABFolderCount )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ ULONG *pCount);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StartPABContactSearch )( 
            IContactFinderCtrl * This,
            /* [in] */ BSTR bstrSearchBy,
            /* [in] */ BSTR bstrStartWith,
            /* [in] */ ULONG lPageNumber);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StartNK2ContactSearch )( 
            IContactFinderCtrl * This,
            /* [in] */ BSTR bstrSearchBy,
            /* [in] */ BSTR bstrStartWith,
            /* [in] */ ULONG lPageNumber);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Abort )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AbortPABSearch )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AbortNK2Search )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StartPABFolderCounting )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StartPABContactCounting )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AbortPABFolderCounting )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AbortPABContactCounting )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_FetchUnique )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ USHORT *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_FetchUnique )( 
            IContactFinderCtrl * This,
            /* [in] */ USHORT newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PABPageSize )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ ULONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_PABPageSize )( 
            IContactFinderCtrl * This,
            /* [in] */ ULONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_NK2PageSize )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ ULONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_NK2PageSize )( 
            IContactFinderCtrl * This,
            /* [in] */ ULONG newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *StartNK2ContactCounting )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AbortNK2ContactCounting )( 
            IContactFinderCtrl * This);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_SubFolderLevel )( 
            IContactFinderCtrl * This,
            /* [in] */ ULONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_EnableLog )( 
            IContactFinderCtrl * This,
            /* [retval][out] */ USHORT *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_EnableLog )( 
            IContactFinderCtrl * This,
            /* [in] */ USHORT newVal);
        
        END_INTERFACE
    } IContactFinderCtrlVtbl;

    interface IContactFinderCtrl
    {
        CONST_VTBL struct IContactFinderCtrlVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IContactFinderCtrl_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IContactFinderCtrl_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IContactFinderCtrl_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IContactFinderCtrl_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IContactFinderCtrl_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IContactFinderCtrl_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IContactFinderCtrl_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IContactFinderCtrl_InitCtrl(This,pInit)	\
    ( (This)->lpVtbl -> InitCtrl(This,pInit) ) 

#define IContactFinderCtrl_GetPABContactsCount(This,pCount)	\
    ( (This)->lpVtbl -> GetPABContactsCount(This,pCount) ) 

#define IContactFinderCtrl_GetNK2ContactsCount(This,pCount)	\
    ( (This)->lpVtbl -> GetNK2ContactsCount(This,pCount) ) 

#define IContactFinderCtrl_GetPABFolderCount(This,pCount)	\
    ( (This)->lpVtbl -> GetPABFolderCount(This,pCount) ) 

#define IContactFinderCtrl_StartPABContactSearch(This,bstrSearchBy,bstrStartWith,lPageNumber)	\
    ( (This)->lpVtbl -> StartPABContactSearch(This,bstrSearchBy,bstrStartWith,lPageNumber) ) 

#define IContactFinderCtrl_StartNK2ContactSearch(This,bstrSearchBy,bstrStartWith,lPageNumber)	\
    ( (This)->lpVtbl -> StartNK2ContactSearch(This,bstrSearchBy,bstrStartWith,lPageNumber) ) 

#define IContactFinderCtrl_Abort(This)	\
    ( (This)->lpVtbl -> Abort(This) ) 

#define IContactFinderCtrl_AbortPABSearch(This)	\
    ( (This)->lpVtbl -> AbortPABSearch(This) ) 

#define IContactFinderCtrl_AbortNK2Search(This)	\
    ( (This)->lpVtbl -> AbortNK2Search(This) ) 

#define IContactFinderCtrl_StartPABFolderCounting(This)	\
    ( (This)->lpVtbl -> StartPABFolderCounting(This) ) 

#define IContactFinderCtrl_StartPABContactCounting(This)	\
    ( (This)->lpVtbl -> StartPABContactCounting(This) ) 

#define IContactFinderCtrl_AbortPABFolderCounting(This)	\
    ( (This)->lpVtbl -> AbortPABFolderCounting(This) ) 

#define IContactFinderCtrl_AbortPABContactCounting(This)	\
    ( (This)->lpVtbl -> AbortPABContactCounting(This) ) 

#define IContactFinderCtrl_get_FetchUnique(This,pVal)	\
    ( (This)->lpVtbl -> get_FetchUnique(This,pVal) ) 

#define IContactFinderCtrl_put_FetchUnique(This,newVal)	\
    ( (This)->lpVtbl -> put_FetchUnique(This,newVal) ) 

#define IContactFinderCtrl_get_PABPageSize(This,pVal)	\
    ( (This)->lpVtbl -> get_PABPageSize(This,pVal) ) 

#define IContactFinderCtrl_put_PABPageSize(This,newVal)	\
    ( (This)->lpVtbl -> put_PABPageSize(This,newVal) ) 

#define IContactFinderCtrl_get_NK2PageSize(This,pVal)	\
    ( (This)->lpVtbl -> get_NK2PageSize(This,pVal) ) 

#define IContactFinderCtrl_put_NK2PageSize(This,newVal)	\
    ( (This)->lpVtbl -> put_NK2PageSize(This,newVal) ) 

#define IContactFinderCtrl_StartNK2ContactCounting(This)	\
    ( (This)->lpVtbl -> StartNK2ContactCounting(This) ) 

#define IContactFinderCtrl_AbortNK2ContactCounting(This)	\
    ( (This)->lpVtbl -> AbortNK2ContactCounting(This) ) 

#define IContactFinderCtrl_put_SubFolderLevel(This,newVal)	\
    ( (This)->lpVtbl -> put_SubFolderLevel(This,newVal) ) 

#define IContactFinderCtrl_get_EnableLog(This,pVal)	\
    ( (This)->lpVtbl -> get_EnableLog(This,pVal) ) 

#define IContactFinderCtrl_put_EnableLog(This,newVal)	\
    ( (This)->lpVtbl -> put_EnableLog(This,newVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IContactFinderCtrl_INTERFACE_DEFINED__ */



#ifndef __TRContactFinderLib_LIBRARY_DEFINED__
#define __TRContactFinderLib_LIBRARY_DEFINED__

/* library TRContactFinderLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_TRContactFinderLib;

#ifndef ___IContactFinderCtrlEvents_DISPINTERFACE_DEFINED__
#define ___IContactFinderCtrlEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IContactFinderCtrlEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IContactFinderCtrlEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("4367F023-3BD7-4e1c-87D2-327DBFF7DF34")
    _IContactFinderCtrlEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IContactFinderCtrlEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IContactFinderCtrlEvents * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IContactFinderCtrlEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IContactFinderCtrlEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IContactFinderCtrlEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IContactFinderCtrlEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IContactFinderCtrlEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IContactFinderCtrlEvents * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _IContactFinderCtrlEventsVtbl;

    interface _IContactFinderCtrlEvents
    {
        CONST_VTBL struct _IContactFinderCtrlEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IContactFinderCtrlEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IContactFinderCtrlEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IContactFinderCtrlEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IContactFinderCtrlEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IContactFinderCtrlEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IContactFinderCtrlEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IContactFinderCtrlEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IContactFinderCtrlEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_ContactFinderCtrl;

#ifdef __cplusplus

class DECLSPEC_UUID("9A4AF171-92B3-4339-B00F-C8FF4A7CD12E")
ContactFinderCtrl;
#endif
#endif /* __TRContactFinderLib_LIBRARY_DEFINED__ */

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


