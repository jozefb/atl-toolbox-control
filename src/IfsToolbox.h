/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Mon Oct 03 22:43:45 2005
 */
/* Compiler settings for D:\Dokumenty\vsproj\Pdb\IfsToolbox\IfsToolbox.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __IfsToolbox_h__
#define __IfsToolbox_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IToolBoxTabs_FWD_DEFINED__
#define __IToolBoxTabs_FWD_DEFINED__
typedef interface IToolBoxTabs IToolBoxTabs;
#endif 	/* __IToolBoxTabs_FWD_DEFINED__ */


#ifndef __IToolBoxTab_FWD_DEFINED__
#define __IToolBoxTab_FWD_DEFINED__
typedef interface IToolBoxTab IToolBoxTab;
#endif 	/* __IToolBoxTab_FWD_DEFINED__ */


#ifndef __IToolBoxItems_FWD_DEFINED__
#define __IToolBoxItems_FWD_DEFINED__
typedef interface IToolBoxItems IToolBoxItems;
#endif 	/* __IToolBoxItems_FWD_DEFINED__ */


#ifndef __IToolBoxItem_FWD_DEFINED__
#define __IToolBoxItem_FWD_DEFINED__
typedef interface IToolBoxItem IToolBoxItem;
#endif 	/* __IToolBoxItem_FWD_DEFINED__ */


#ifndef __IToolbox_FWD_DEFINED__
#define __IToolbox_FWD_DEFINED__
typedef interface IToolbox IToolbox;
#endif 	/* __IToolbox_FWD_DEFINED__ */


#ifndef __IToolboxEvents_FWD_DEFINED__
#define __IToolboxEvents_FWD_DEFINED__
typedef interface IToolboxEvents IToolboxEvents;
#endif 	/* __IToolboxEvents_FWD_DEFINED__ */


#ifndef __Toolbox_FWD_DEFINED__
#define __Toolbox_FWD_DEFINED__

#ifdef __cplusplus
typedef class Toolbox Toolbox;
#else
typedef struct Toolbox Toolbox;
#endif /* __cplusplus */

#endif 	/* __Toolbox_FWD_DEFINED__ */


#ifndef __ToolBoxTabs_FWD_DEFINED__
#define __ToolBoxTabs_FWD_DEFINED__

#ifdef __cplusplus
typedef class ToolBoxTabs ToolBoxTabs;
#else
typedef struct ToolBoxTabs ToolBoxTabs;
#endif /* __cplusplus */

#endif 	/* __ToolBoxTabs_FWD_DEFINED__ */


#ifndef __ToolBoxTab_FWD_DEFINED__
#define __ToolBoxTab_FWD_DEFINED__

#ifdef __cplusplus
typedef class ToolBoxTab ToolBoxTab;
#else
typedef struct ToolBoxTab ToolBoxTab;
#endif /* __cplusplus */

#endif 	/* __ToolBoxTab_FWD_DEFINED__ */


#ifndef __ToolBoxItem_FWD_DEFINED__
#define __ToolBoxItem_FWD_DEFINED__

#ifdef __cplusplus
typedef class ToolBoxItem ToolBoxItem;
#else
typedef struct ToolBoxItem ToolBoxItem;
#endif /* __cplusplus */

#endif 	/* __ToolBoxItem_FWD_DEFINED__ */


#ifndef __ToolBoxItems_FWD_DEFINED__
#define __ToolBoxItems_FWD_DEFINED__

#ifdef __cplusplus
typedef class ToolBoxItems ToolBoxItems;
#else
typedef struct ToolBoxItems ToolBoxItems;
#endif /* __cplusplus */

#endif 	/* __ToolBoxItems_FWD_DEFINED__ */


#ifndef __PPToolbox_FWD_DEFINED__
#define __PPToolbox_FWD_DEFINED__

#ifdef __cplusplus
typedef class PPToolbox PPToolbox;
#else
typedef struct PPToolbox PPToolbox;
#endif /* __cplusplus */

#endif 	/* __PPToolbox_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 


#ifndef __IFSTOOLBOXLib_LIBRARY_DEFINED__
#define __IFSTOOLBOXLib_LIBRARY_DEFINED__

/* library IFSTOOLBOXLib */
/* [helpstring][version][uuid] */ 





typedef /* [v1_enum] */ 
enum ToolBoxItemFormat
    {	ToolBoxItemFormatText	= 0,
	ToolBoxItemFormatHTML	= 1,
	ToolBoxItemFormatGUID	= 2,
	ToolBoxItemFormatPointer	= 3
    }	ToolBoxItemFormat;


EXTERN_C const IID LIBID_IFSTOOLBOXLib;

#ifndef __IToolBoxTabs_INTERFACE_DEFINED__
#define __IToolBoxTabs_INTERFACE_DEFINED__

/* interface IToolBoxTabs */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IToolBoxTabs;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A7CEA18F-93F9-4CBE-90A1-8CFEF82BE11A")
    IToolBoxTabs : public IDispatch
    {
    public:
        virtual /* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IToolbox __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DTE( 
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            VARIANT index,
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR bstrName,
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pToolBoxTab) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IToolBoxTabsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IToolBoxTabs __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IToolBoxTabs __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Parent )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [retval][out] */ IToolbox __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DTE )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Item )( 
            IToolBoxTabs __RPC_FAR * This,
            VARIANT index,
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            IToolBoxTabs __RPC_FAR * This,
            /* [in] */ BSTR bstrName,
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pToolBoxTab);
        
        END_INTERFACE
    } IToolBoxTabsVtbl;

    interface IToolBoxTabs
    {
        CONST_VTBL struct IToolBoxTabsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IToolBoxTabs_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IToolBoxTabs_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IToolBoxTabs_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IToolBoxTabs_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IToolBoxTabs_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IToolBoxTabs_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IToolBoxTabs_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IToolBoxTabs_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define IToolBoxTabs_get_Parent(This,pVal)	\
    (This)->lpVtbl -> get_Parent(This,pVal)

#define IToolBoxTabs_get_DTE(This,pVal)	\
    (This)->lpVtbl -> get_DTE(This,pVal)

#define IToolBoxTabs_get_Item(This,index,pVal)	\
    (This)->lpVtbl -> get_Item(This,index,pVal)

#define IToolBoxTabs_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define IToolBoxTabs_Add(This,bstrName,pToolBoxTab)	\
    (This)->lpVtbl -> Add(This,bstrName,pToolBoxTab)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTabs_get__NewEnum_Proxy( 
    IToolBoxTabs __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTabs_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTabs_get_Parent_Proxy( 
    IToolBoxTabs __RPC_FAR * This,
    /* [retval][out] */ IToolbox __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxTabs_get_Parent_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTabs_get_DTE_Proxy( 
    IToolBoxTabs __RPC_FAR * This,
    /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTabs_get_DTE_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTabs_get_Item_Proxy( 
    IToolBoxTabs __RPC_FAR * This,
    VARIANT index,
    /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxTabs_get_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTabs_get_Count_Proxy( 
    IToolBoxTabs __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTabs_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IToolBoxTabs_Add_Proxy( 
    IToolBoxTabs __RPC_FAR * This,
    /* [in] */ BSTR bstrName,
    /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pToolBoxTab);


void __RPC_STUB IToolBoxTabs_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IToolBoxTabs_INTERFACE_DEFINED__ */


#ifndef __IToolBoxTab_INTERFACE_DEFINED__
#define __IToolBoxTab_INTERFACE_DEFINED__

/* interface IToolBoxTab */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IToolBoxTab;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0141D84E-2565-4DE1-AF95-BF70273501CB")
    IToolBoxTab : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Guid( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Guid( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Delete( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Activate( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Collection( 
            /* [retval][out] */ IToolBoxTabs __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ListView( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ListView( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DTE( 
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ToolBoxItems( 
            /* [retval][out] */ IToolBoxItems __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Deleteable( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Deleteable( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Hidden( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Hidden( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ID( 
            /* [retval][out] */ int __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IToolBoxTabVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IToolBoxTab __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IToolBoxTab __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IToolBoxTab __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Name )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Guid )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Guid )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Delete )( 
            IToolBoxTab __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Activate )( 
            IToolBoxTab __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Collection )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ IToolBoxTabs __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ListView )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ListView )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DTE )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ToolBoxItems )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ IToolBoxItems __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Deleteable )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Deleteable )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Hidden )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Hidden )( 
            IToolBoxTab __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ID )( 
            IToolBoxTab __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *pVal);
        
        END_INTERFACE
    } IToolBoxTabVtbl;

    interface IToolBoxTab
    {
        CONST_VTBL struct IToolBoxTabVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IToolBoxTab_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IToolBoxTab_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IToolBoxTab_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IToolBoxTab_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IToolBoxTab_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IToolBoxTab_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IToolBoxTab_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IToolBoxTab_get_Name(This,pVal)	\
    (This)->lpVtbl -> get_Name(This,pVal)

#define IToolBoxTab_put_Name(This,newVal)	\
    (This)->lpVtbl -> put_Name(This,newVal)

#define IToolBoxTab_get_Guid(This,pVal)	\
    (This)->lpVtbl -> get_Guid(This,pVal)

#define IToolBoxTab_put_Guid(This,newVal)	\
    (This)->lpVtbl -> put_Guid(This,newVal)

#define IToolBoxTab_Delete(This)	\
    (This)->lpVtbl -> Delete(This)

#define IToolBoxTab_Activate(This)	\
    (This)->lpVtbl -> Activate(This)

#define IToolBoxTab_get_Collection(This,pVal)	\
    (This)->lpVtbl -> get_Collection(This,pVal)

#define IToolBoxTab_get_ListView(This,pVal)	\
    (This)->lpVtbl -> get_ListView(This,pVal)

#define IToolBoxTab_put_ListView(This,newVal)	\
    (This)->lpVtbl -> put_ListView(This,newVal)

#define IToolBoxTab_get_DTE(This,pVal)	\
    (This)->lpVtbl -> get_DTE(This,pVal)

#define IToolBoxTab_get_ToolBoxItems(This,pVal)	\
    (This)->lpVtbl -> get_ToolBoxItems(This,pVal)

#define IToolBoxTab_get_Deleteable(This,pVal)	\
    (This)->lpVtbl -> get_Deleteable(This,pVal)

#define IToolBoxTab_put_Deleteable(This,newVal)	\
    (This)->lpVtbl -> put_Deleteable(This,newVal)

#define IToolBoxTab_get_Hidden(This,pVal)	\
    (This)->lpVtbl -> get_Hidden(This,pVal)

#define IToolBoxTab_put_Hidden(This,newVal)	\
    (This)->lpVtbl -> put_Hidden(This,newVal)

#define IToolBoxTab_get_ID(This,pVal)	\
    (This)->lpVtbl -> get_ID(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_Name_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_put_Name_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IToolBoxTab_put_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_Guid_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_Guid_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_put_Guid_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IToolBoxTab_put_Guid_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_Delete_Proxy( 
    IToolBoxTab __RPC_FAR * This);


void __RPC_STUB IToolBoxTab_Delete_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_Activate_Proxy( 
    IToolBoxTab __RPC_FAR * This);


void __RPC_STUB IToolBoxTab_Activate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_Collection_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ IToolBoxTabs __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_Collection_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_ListView_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_ListView_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_put_ListView_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IToolBoxTab_put_ListView_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_DTE_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_DTE_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_ToolBoxItems_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ IToolBoxItems __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_ToolBoxItems_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_Deleteable_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_Deleteable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_put_Deleteable_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IToolBoxTab_put_Deleteable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_Hidden_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_Hidden_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_put_Hidden_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IToolBoxTab_put_Hidden_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxTab_get_ID_Proxy( 
    IToolBoxTab __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *pVal);


void __RPC_STUB IToolBoxTab_get_ID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IToolBoxTab_INTERFACE_DEFINED__ */


#ifndef __IToolBoxItems_INTERFACE_DEFINED__
#define __IToolBoxItems_INTERFACE_DEFINED__

/* interface IToolBoxItems */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IToolBoxItems;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("20C96C4B-34A5-4BC5-AE4D-3CC97B78C8FC")
    IToolBoxItems : public IDispatch
    {
    public:
        virtual /* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            VARIANT index,
            /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SelectedItem( 
            /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR Name,
            /* [in] */ VARIANT Data,
            /* [defaultvalue][optional][in] */ ToolBoxItemFormat Format,
            /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *ppItem) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DTE( 
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IToolBoxItemsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IToolBoxItems __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IToolBoxItems __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IToolBoxItems __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IToolBoxItems __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IToolBoxItems __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IToolBoxItems __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IToolBoxItems __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get__NewEnum )( 
            IToolBoxItems __RPC_FAR * This,
            /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Item )( 
            IToolBoxItems __RPC_FAR * This,
            VARIANT index,
            /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            IToolBoxItems __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SelectedItem )( 
            IToolBoxItems __RPC_FAR * This,
            /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Parent )( 
            IToolBoxItems __RPC_FAR * This,
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            IToolBoxItems __RPC_FAR * This,
            /* [in] */ BSTR Name,
            /* [in] */ VARIANT Data,
            /* [defaultvalue][optional][in] */ ToolBoxItemFormat Format,
            /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *ppItem);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DTE )( 
            IToolBoxItems __RPC_FAR * This,
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);
        
        END_INTERFACE
    } IToolBoxItemsVtbl;

    interface IToolBoxItems
    {
        CONST_VTBL struct IToolBoxItemsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IToolBoxItems_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IToolBoxItems_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IToolBoxItems_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IToolBoxItems_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IToolBoxItems_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IToolBoxItems_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IToolBoxItems_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IToolBoxItems_get__NewEnum(This,pVal)	\
    (This)->lpVtbl -> get__NewEnum(This,pVal)

#define IToolBoxItems_get_Item(This,index,pVal)	\
    (This)->lpVtbl -> get_Item(This,index,pVal)

#define IToolBoxItems_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define IToolBoxItems_get_SelectedItem(This,pVal)	\
    (This)->lpVtbl -> get_SelectedItem(This,pVal)

#define IToolBoxItems_get_Parent(This,pVal)	\
    (This)->lpVtbl -> get_Parent(This,pVal)

#define IToolBoxItems_Add(This,Name,Data,Format,ppItem)	\
    (This)->lpVtbl -> Add(This,Name,Data,Format,ppItem)

#define IToolBoxItems_get_DTE(This,pVal)	\
    (This)->lpVtbl -> get_DTE(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItems_get__NewEnum_Proxy( 
    IToolBoxItems __RPC_FAR * This,
    /* [retval][out] */ LPUNKNOWN __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItems_get__NewEnum_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItems_get_Item_Proxy( 
    IToolBoxItems __RPC_FAR * This,
    VARIANT index,
    /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxItems_get_Item_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItems_get_Count_Proxy( 
    IToolBoxItems __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItems_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItems_get_SelectedItem_Proxy( 
    IToolBoxItems __RPC_FAR * This,
    /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxItems_get_SelectedItem_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItems_get_Parent_Proxy( 
    IToolBoxItems __RPC_FAR * This,
    /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxItems_get_Parent_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IToolBoxItems_Add_Proxy( 
    IToolBoxItems __RPC_FAR * This,
    /* [in] */ BSTR Name,
    /* [in] */ VARIANT Data,
    /* [defaultvalue][optional][in] */ ToolBoxItemFormat Format,
    /* [retval][out] */ IToolBoxItem __RPC_FAR *__RPC_FAR *ppItem);


void __RPC_STUB IToolBoxItems_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItems_get_DTE_Proxy( 
    IToolBoxItems __RPC_FAR * This,
    /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItems_get_DTE_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IToolBoxItems_INTERFACE_DEFINED__ */


#ifndef __IToolBoxItem_INTERFACE_DEFINED__
#define __IToolBoxItem_INTERFACE_DEFINED__

/* interface IToolBoxItem */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IToolBoxItem;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("AD281E04-D150-4FD1-A74D-2F7C86658742")
    IToolBoxItem : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Guid( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Guid( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Delete( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Select( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Collection( 
            /* [retval][out] */ IToolBoxItems __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DTE( 
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Format( 
            /* [retval][out] */ ToolBoxItemFormat __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Format( 
            /* [in] */ ToolBoxItemFormat newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Data( 
            /* [retval][out] */ VARIANT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Data( 
            /* [in] */ VARIANT newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Description( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Description( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Hidden( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Hidden( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Deleteable( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Deleteable( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ID( 
            /* [retval][out] */ int __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IToolBoxItemVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IToolBoxItem __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IToolBoxItem __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IToolBoxItem __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Name )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Name )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Guid )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Guid )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Delete )( 
            IToolBoxItem __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Select )( 
            IToolBoxItem __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Collection )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ IToolBoxItems __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DTE )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Format )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ ToolBoxItemFormat __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Format )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ ToolBoxItemFormat newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Data )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ VARIANT __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Data )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ VARIANT newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Description )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Description )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Hidden )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Hidden )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Deleteable )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Deleteable )( 
            IToolBoxItem __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ID )( 
            IToolBoxItem __RPC_FAR * This,
            /* [retval][out] */ int __RPC_FAR *pVal);
        
        END_INTERFACE
    } IToolBoxItemVtbl;

    interface IToolBoxItem
    {
        CONST_VTBL struct IToolBoxItemVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IToolBoxItem_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IToolBoxItem_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IToolBoxItem_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IToolBoxItem_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IToolBoxItem_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IToolBoxItem_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IToolBoxItem_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IToolBoxItem_get_Name(This,pVal)	\
    (This)->lpVtbl -> get_Name(This,pVal)

#define IToolBoxItem_put_Name(This,newVal)	\
    (This)->lpVtbl -> put_Name(This,newVal)

#define IToolBoxItem_get_Guid(This,pVal)	\
    (This)->lpVtbl -> get_Guid(This,pVal)

#define IToolBoxItem_put_Guid(This,newVal)	\
    (This)->lpVtbl -> put_Guid(This,newVal)

#define IToolBoxItem_Delete(This)	\
    (This)->lpVtbl -> Delete(This)

#define IToolBoxItem_Select(This)	\
    (This)->lpVtbl -> Select(This)

#define IToolBoxItem_get_Collection(This,pVal)	\
    (This)->lpVtbl -> get_Collection(This,pVal)

#define IToolBoxItem_get_DTE(This,pVal)	\
    (This)->lpVtbl -> get_DTE(This,pVal)

#define IToolBoxItem_get_Format(This,pVal)	\
    (This)->lpVtbl -> get_Format(This,pVal)

#define IToolBoxItem_put_Format(This,newVal)	\
    (This)->lpVtbl -> put_Format(This,newVal)

#define IToolBoxItem_get_Data(This,pVal)	\
    (This)->lpVtbl -> get_Data(This,pVal)

#define IToolBoxItem_put_Data(This,newVal)	\
    (This)->lpVtbl -> put_Data(This,newVal)

#define IToolBoxItem_get_Description(This,pVal)	\
    (This)->lpVtbl -> get_Description(This,pVal)

#define IToolBoxItem_put_Description(This,newVal)	\
    (This)->lpVtbl -> put_Description(This,newVal)

#define IToolBoxItem_get_Hidden(This,pVal)	\
    (This)->lpVtbl -> get_Hidden(This,pVal)

#define IToolBoxItem_put_Hidden(This,newVal)	\
    (This)->lpVtbl -> put_Hidden(This,newVal)

#define IToolBoxItem_get_Deleteable(This,pVal)	\
    (This)->lpVtbl -> get_Deleteable(This,pVal)

#define IToolBoxItem_put_Deleteable(This,newVal)	\
    (This)->lpVtbl -> put_Deleteable(This,newVal)

#define IToolBoxItem_get_ID(This,pVal)	\
    (This)->lpVtbl -> get_ID(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Name_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_put_Name_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IToolBoxItem_put_Name_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Guid_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Guid_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_put_Guid_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IToolBoxItem_put_Guid_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_Delete_Proxy( 
    IToolBoxItem __RPC_FAR * This);


void __RPC_STUB IToolBoxItem_Delete_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_Select_Proxy( 
    IToolBoxItem __RPC_FAR * This);


void __RPC_STUB IToolBoxItem_Select_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Collection_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ IToolBoxItems __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Collection_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_DTE_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_DTE_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Format_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ ToolBoxItemFormat __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Format_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_put_Format_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [in] */ ToolBoxItemFormat newVal);


void __RPC_STUB IToolBoxItem_put_Format_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Data_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ VARIANT __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Data_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_put_Data_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [in] */ VARIANT newVal);


void __RPC_STUB IToolBoxItem_put_Data_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Description_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Description_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_put_Description_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IToolBoxItem_put_Description_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Hidden_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Hidden_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_put_Hidden_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IToolBoxItem_put_Hidden_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_Deleteable_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_Deleteable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_put_Deleteable_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IToolBoxItem_put_Deleteable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolBoxItem_get_ID_Proxy( 
    IToolBoxItem __RPC_FAR * This,
    /* [retval][out] */ int __RPC_FAR *pVal);


void __RPC_STUB IToolBoxItem_get_ID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IToolBoxItem_INTERFACE_DEFINED__ */


#ifndef __IToolbox_INTERFACE_DEFINED__
#define __IToolbox_INTERFACE_DEFINED__

/* interface IToolbox */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IToolbox;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("4A5854D2-EEC7-415F-B7C9-C20FB1A899E2")
    IToolbox : public IDispatch
    {
    public:
        virtual /* [id][propputref] */ HRESULT STDMETHODCALLTYPE putref_Font( 
            /* [in] */ IFontDisp __RPC_FAR *pFont) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Font( 
            /* [in] */ IFontDisp __RPC_FAR *pFont) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Font( 
            /* [retval][out] */ IFontDisp __RPC_FAR *__RPC_FAR *ppFont) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Enabled( 
            /* [in] */ VARIANT_BOOL vbool) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Enabled( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbool) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Window( 
            /* [retval][out] */ long __RPC_FAR *phwnd) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ToolBoxTabs( 
            /* [retval][out] */ IToolBoxTabs __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ActiveTab( 
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DTE( 
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_DTE( 
            /* [in] */ LPDISPATCH newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ShowPopupMenu( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ShowPopupMenu( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PMICustomize( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_PMICustomize( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IToolboxVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IToolbox __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IToolbox __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IToolbox __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id][propputref] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *putref_Font )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ IFontDisp __RPC_FAR *pFont);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Font )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ IFontDisp __RPC_FAR *pFont);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Font )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ IFontDisp __RPC_FAR *__RPC_FAR *ppFont);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Enabled )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL vbool);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Enabled )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbool);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Window )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *phwnd);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ToolBoxTabs )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ IToolBoxTabs __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ActiveTab )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DTE )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_DTE )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ LPDISPATCH newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ShowPopupMenu )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ShowPopupMenu )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_PMICustomize )( 
            IToolbox __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_PMICustomize )( 
            IToolbox __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        END_INTERFACE
    } IToolboxVtbl;

    interface IToolbox
    {
        CONST_VTBL struct IToolboxVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IToolbox_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IToolbox_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IToolbox_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IToolbox_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IToolbox_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IToolbox_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IToolbox_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IToolbox_putref_Font(This,pFont)	\
    (This)->lpVtbl -> putref_Font(This,pFont)

#define IToolbox_put_Font(This,pFont)	\
    (This)->lpVtbl -> put_Font(This,pFont)

#define IToolbox_get_Font(This,ppFont)	\
    (This)->lpVtbl -> get_Font(This,ppFont)

#define IToolbox_put_Enabled(This,vbool)	\
    (This)->lpVtbl -> put_Enabled(This,vbool)

#define IToolbox_get_Enabled(This,pbool)	\
    (This)->lpVtbl -> get_Enabled(This,pbool)

#define IToolbox_get_Window(This,phwnd)	\
    (This)->lpVtbl -> get_Window(This,phwnd)

#define IToolbox_get_ToolBoxTabs(This,pVal)	\
    (This)->lpVtbl -> get_ToolBoxTabs(This,pVal)

#define IToolbox_get_ActiveTab(This,pVal)	\
    (This)->lpVtbl -> get_ActiveTab(This,pVal)

#define IToolbox_get_DTE(This,pVal)	\
    (This)->lpVtbl -> get_DTE(This,pVal)

#define IToolbox_put_DTE(This,newVal)	\
    (This)->lpVtbl -> put_DTE(This,newVal)

#define IToolbox_get_ShowPopupMenu(This,pVal)	\
    (This)->lpVtbl -> get_ShowPopupMenu(This,pVal)

#define IToolbox_put_ShowPopupMenu(This,newVal)	\
    (This)->lpVtbl -> put_ShowPopupMenu(This,newVal)

#define IToolbox_get_PMICustomize(This,pVal)	\
    (This)->lpVtbl -> get_PMICustomize(This,pVal)

#define IToolbox_put_PMICustomize(This,newVal)	\
    (This)->lpVtbl -> put_PMICustomize(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id][propputref] */ HRESULT STDMETHODCALLTYPE IToolbox_putref_Font_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [in] */ IFontDisp __RPC_FAR *pFont);


void __RPC_STUB IToolbox_putref_Font_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IToolbox_put_Font_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [in] */ IFontDisp __RPC_FAR *pFont);


void __RPC_STUB IToolbox_put_Font_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_Font_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ IFontDisp __RPC_FAR *__RPC_FAR *ppFont);


void __RPC_STUB IToolbox_get_Font_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IToolbox_put_Enabled_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL vbool);


void __RPC_STUB IToolbox_put_Enabled_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_Enabled_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pbool);


void __RPC_STUB IToolbox_get_Enabled_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_Window_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *phwnd);


void __RPC_STUB IToolbox_get_Window_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_ToolBoxTabs_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ IToolBoxTabs __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolbox_get_ToolBoxTabs_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_ActiveTab_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ IToolBoxTab __RPC_FAR *__RPC_FAR *pVal);


void __RPC_STUB IToolbox_get_ActiveTab_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_DTE_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ LPDISPATCH __RPC_FAR *pVal);


void __RPC_STUB IToolbox_get_DTE_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolbox_put_DTE_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [in] */ LPDISPATCH newVal);


void __RPC_STUB IToolbox_put_DTE_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_ShowPopupMenu_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IToolbox_get_ShowPopupMenu_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolbox_put_ShowPopupMenu_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IToolbox_put_ShowPopupMenu_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IToolbox_get_PMICustomize_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IToolbox_get_PMICustomize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IToolbox_put_PMICustomize_Proxy( 
    IToolbox __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IToolbox_put_PMICustomize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IToolbox_INTERFACE_DEFINED__ */


#ifndef __IToolboxEvents_DISPINTERFACE_DEFINED__
#define __IToolboxEvents_DISPINTERFACE_DEFINED__

/* dispinterface IToolboxEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IToolboxEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("E9CA9CB0-1CA0-4C7A-9CE7-DA2EDC1D40DD")
    IToolboxEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IToolboxEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IToolboxEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IToolboxEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IToolboxEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IToolboxEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IToolboxEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IToolboxEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IToolboxEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } IToolboxEventsVtbl;

    interface IToolboxEvents
    {
        CONST_VTBL struct IToolboxEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IToolboxEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IToolboxEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IToolboxEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IToolboxEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IToolboxEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IToolboxEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IToolboxEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IToolboxEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Toolbox;

#ifdef __cplusplus

class DECLSPEC_UUID("CA65189C-6495-4C3B-9762-04476D998250")
Toolbox;
#endif

EXTERN_C const CLSID CLSID_ToolBoxTabs;

#ifdef __cplusplus

class DECLSPEC_UUID("8AD3C7A9-6494-44CE-89C5-50CDE5579D55")
ToolBoxTabs;
#endif

EXTERN_C const CLSID CLSID_ToolBoxTab;

#ifdef __cplusplus

class DECLSPEC_UUID("466BBEE5-1B6C-4C27-AE0D-1F9F2FDC1478")
ToolBoxTab;
#endif

EXTERN_C const CLSID CLSID_ToolBoxItem;

#ifdef __cplusplus

class DECLSPEC_UUID("E0EED2DB-2B17-4470-8BDD-FCBBA583A7B2")
ToolBoxItem;
#endif

EXTERN_C const CLSID CLSID_ToolBoxItems;

#ifdef __cplusplus

class DECLSPEC_UUID("C1AD471F-724F-48EB-8264-BF9E827FC547")
ToolBoxItems;
#endif

EXTERN_C const CLSID CLSID_PPToolbox;

#ifdef __cplusplus

class DECLSPEC_UUID("CB212624-90E3-44B5-9E1A-6DE7CC814A55")
PPToolbox;
#endif
#endif /* __IFSTOOLBOXLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
