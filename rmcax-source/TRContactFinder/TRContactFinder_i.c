

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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

MIDL_DEFINE_GUID(IID, IID_IContactFinderCtrl,0xCE2074C6,0x86AB,0x4f46,0x95,0xF8,0x9B,0x4B,0x20,0x88,0xE3,0x73);


MIDL_DEFINE_GUID(IID, LIBID_TRContactFinderLib,0x8C1D9D24,0x38CE,0x4a3f,0x80,0x6A,0xCB,0x17,0xEE,0xBA,0xE9,0xA0);


MIDL_DEFINE_GUID(IID, DIID__IContactFinderCtrlEvents,0x4367F023,0x3BD7,0x4e1c,0x87,0xD2,0x32,0x7D,0xBF,0xF7,0xDF,0x34);


MIDL_DEFINE_GUID(CLSID, CLSID_ContactFinderCtrl,0x9A4AF171,0x92B3,0x4339,0xB0,0x0F,0xC8,0xFF,0x4A,0x7C,0xD1,0x2E);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



