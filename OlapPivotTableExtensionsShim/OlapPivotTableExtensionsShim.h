

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Sun May 30 13:20:20 2010
 */
/* Compiler settings for .\OlapPivotTableExtensionsShim.idl:
    Oicf, W1, Zp8, env=Win64 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
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


#ifndef __OlapPivotTableExtensionsShim_h__
#define __OlapPivotTableExtensionsShim_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ConnectProxy_FWD_DEFINED__
#define __ConnectProxy_FWD_DEFINED__

#ifdef __cplusplus
typedef class ConnectProxy ConnectProxy;
#else
typedef struct ConnectProxy ConnectProxy;
#endif /* __cplusplus */

#endif 	/* __ConnectProxy_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __OlapPivotTableExtensionsShimLib_LIBRARY_DEFINED__
#define __OlapPivotTableExtensionsShimLib_LIBRARY_DEFINED__

/* library OlapPivotTableExtensionsShimLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_OlapPivotTableExtensionsShimLib;

EXTERN_C const CLSID CLSID_ConnectProxy;

#ifdef __cplusplus

class DECLSPEC_UUID("dd16a145-e2f0-40b9-9993-5018ba8b6ff3")
ConnectProxy;
#endif
#endif /* __OlapPivotTableExtensionsShimLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


