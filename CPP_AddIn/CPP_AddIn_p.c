

/* this ALWAYS GENERATED file contains the proxy stub code */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 03:14:07 2038
 */
/* Compiler settings for CPP_AddIn.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#if !defined(_M_IA64) && !defined(_M_AMD64) && !defined(_ARM_)


#if _MSC_VER >= 1200
#pragma warning(push)
#endif

#pragma warning( disable: 4211 )  /* redefine extern to static */
#pragma warning( disable: 4232 )  /* dllimport identity*/
#pragma warning( disable: 4024 )  /* array to pointer mapping*/
#pragma warning( disable: 4152 )  /* function/data pointer conversion in expression */
#pragma warning( disable: 4100 ) /* unreferenced arguments in x86 call */

#pragma optimize("", off ) 

#define USE_STUBLESS_PROXY


/* verify that the <rpcproxy.h> version is high enough to compile this file*/
#ifndef __REDQ_RPCPROXY_H_VERSION__
#define __REQUIRED_RPCPROXY_H_VERSION__ 475
#endif


#include "rpcproxy.h"
#ifndef __RPCPROXY_H_VERSION__
#error this stub requires an updated version of <rpcproxy.h>
#endif /* __RPCPROXY_H_VERSION__ */


#include "CPP_AddIn_i.h"

#define TYPE_FORMAT_STRING_SIZE   3                                 
#define PROC_FORMAT_STRING_SIZE   1                                 
#define EXPR_FORMAT_STRING_SIZE   1                                 
#define TRANSMIT_AS_TABLE_SIZE    0            
#define WIRE_MARSHAL_TABLE_SIZE   0            

typedef struct _CPP_AddIn_MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } CPP_AddIn_MIDL_TYPE_FORMAT_STRING;

typedef struct _CPP_AddIn_MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } CPP_AddIn_MIDL_PROC_FORMAT_STRING;

typedef struct _CPP_AddIn_MIDL_EXPR_FORMAT_STRING
    {
    long          Pad;
    unsigned char  Format[ EXPR_FORMAT_STRING_SIZE ];
    } CPP_AddIn_MIDL_EXPR_FORMAT_STRING;


static const RPC_SYNTAX_IDENTIFIER  _RpcTransferSyntax = 
{{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}};


extern const CPP_AddIn_MIDL_TYPE_FORMAT_STRING CPP_AddIn__MIDL_TypeFormatString;
extern const CPP_AddIn_MIDL_PROC_FORMAT_STRING CPP_AddIn__MIDL_ProcFormatString;
extern const CPP_AddIn_MIDL_EXPR_FORMAT_STRING CPP_AddIn__MIDL_ExprFormatString;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IConnect_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IConnect_ProxyInfo;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IItemWrapper_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IItemWrapper_ProxyInfo;



#if !defined(__RPC_WIN32__)
#error  Invalid build platform for this stub.
#endif
#if !(TARGET_IS_NT60_OR_LATER)
#error You need Windows Vista or later to run this stub because it uses these features:
#error   forced complex structure or array, new range semantics, compiled for Windows Vista.
#error However, your C/C++ compilation flags indicate you intend to run this app on earlier systems.
#error This app will fail with the RPC_X_WRONG_STUB_VERSION error.
#endif


static const CPP_AddIn_MIDL_PROC_FORMAT_STRING CPP_AddIn__MIDL_ProcFormatString =
    {
        0,
        {

			0x0
        }
    };

static const CPP_AddIn_MIDL_TYPE_FORMAT_STRING CPP_AddIn__MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */

			0x0
        }
    };


/* Object interface: IUnknown, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IDispatch, ver. 0.0,
   GUID={0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IConnect, ver. 0.0,
   GUID={0x91e702f2,0xd64d,0x48e3,{0xb5,0xdd,0x9a,0xfa,0xee,0xad,0x3a,0x00}} */

#pragma code_seg(".orpc")
static const unsigned short IConnect_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IConnect_ProxyInfo =
    {
    &Object_StubDesc,
    CPP_AddIn__MIDL_ProcFormatString.Format,
    &IConnect_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IConnect_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    CPP_AddIn__MIDL_ProcFormatString.Format,
    &IConnect_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IConnectProxyVtbl = 
{
    0,
    &IID_IConnect,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IConnect_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IConnectStubVtbl =
{
    &IID_IConnect,
    &IConnect_ServerInfo,
    7,
    &IConnect_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IItemWrapper, ver. 0.0,
   GUID={0x3c8bd09f,0x3ba0,0x4992,{0xbe,0x24,0xb6,0xc1,0x7d,0xc7,0xe7,0x8a}} */

#pragma code_seg(".orpc")
static const unsigned short IItemWrapper_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IItemWrapper_ProxyInfo =
    {
    &Object_StubDesc,
    CPP_AddIn__MIDL_ProcFormatString.Format,
    &IItemWrapper_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IItemWrapper_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    CPP_AddIn__MIDL_ProcFormatString.Format,
    &IItemWrapper_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IItemWrapperProxyVtbl = 
{
    0,
    &IID_IItemWrapper,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IItemWrapper_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IItemWrapperStubVtbl =
{
    &IID_IItemWrapper,
    &IItemWrapper_ServerInfo,
    7,
    &IItemWrapper_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};

static const MIDL_STUB_DESC Object_StubDesc = 
    {
    0,
    NdrOleAllocate,
    NdrOleFree,
    0,
    0,
    0,
    0,
    0,
    CPP_AddIn__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x60001, /* Ndr library version */
    0,
    0x801026e, /* MIDL Version 8.1.622 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    0x1, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    0
    };

const CInterfaceProxyVtbl * const _CPP_AddIn_ProxyVtblList[] = 
{
    ( CInterfaceProxyVtbl *) &_IItemWrapperProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IConnectProxyVtbl,
    0
};

const CInterfaceStubVtbl * const _CPP_AddIn_StubVtblList[] = 
{
    ( CInterfaceStubVtbl *) &_IItemWrapperStubVtbl,
    ( CInterfaceStubVtbl *) &_IConnectStubVtbl,
    0
};

PCInterfaceName const _CPP_AddIn_InterfaceNamesList[] = 
{
    "IItemWrapper",
    "IConnect",
    0
};

const IID *  const _CPP_AddIn_BaseIIDList[] = 
{
    &IID_IDispatch,
    &IID_IDispatch,
    0
};


#define _CPP_AddIn_CHECK_IID(n)	IID_GENERIC_CHECK_IID( _CPP_AddIn, pIID, n)

int __stdcall _CPP_AddIn_IID_Lookup( const IID * pIID, int * pIndex )
{
    IID_BS_LOOKUP_SETUP

    IID_BS_LOOKUP_INITIAL_TEST( _CPP_AddIn, 2, 1 )
    IID_BS_LOOKUP_RETURN_RESULT( _CPP_AddIn, 2, *pIndex )
    
}

const ExtendedProxyFileInfo CPP_AddIn_ProxyFileInfo = 
{
    (PCInterfaceProxyVtblList *) & _CPP_AddIn_ProxyVtblList,
    (PCInterfaceStubVtblList *) & _CPP_AddIn_StubVtblList,
    (const PCInterfaceName * ) & _CPP_AddIn_InterfaceNamesList,
    (const IID ** ) & _CPP_AddIn_BaseIIDList,
    & _CPP_AddIn_IID_Lookup, 
    2,
    2,
    0, /* table of [async_uuid] interfaces */
    0, /* Filler1 */
    0, /* Filler2 */
    0  /* Filler3 */
};
#pragma optimize("", on )
#if _MSC_VER >= 1200
#pragma warning(pop)
#endif


#endif /* !defined(_M_IA64) && !defined(_M_AMD64) && !defined(_ARM_) */

