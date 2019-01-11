

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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
        EXTERN_C __declspec(selectany) const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif // !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IConnect,0x91e702f2,0xd64d,0x48e3,0xb5,0xdd,0x9a,0xfa,0xee,0xad,0x3a,0x00);


MIDL_DEFINE_GUID(IID, IID_IItemWrapper,0x3c8bd09f,0x3ba0,0x4992,0xbe,0x24,0xb6,0xc1,0x7d,0xc7,0xe7,0x8a);


MIDL_DEFINE_GUID(IID, LIBID_CPP_AddInLib,0xf3eda8a5,0x30af,0x4dc3,0x87,0x5b,0x07,0x51,0x67,0xf2,0x1b,0x10);


MIDL_DEFINE_GUID(CLSID, CLSID_Connect,0x69eb1c58,0x6e7e,0x420e,0xb3,0x39,0x62,0x71,0xf8,0x04,0xe9,0x66);


MIDL_DEFINE_GUID(CLSID, CLSID_ItemWrapper,0xda1194fa,0x238a,0x4f0e,0x8f,0x55,0xd0,0x1c,0x37,0xb3,0x97,0x87);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



