// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__8A41C9D2_09C3_422B_B832_438E87D7ED55__INCLUDED_)
#define AFX_STDAFX_H__8A41C9D2_09C3_422B_B832_438E87D7ED55__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#ifndef _WIN32_IE
#define _WIN32_IE 0x0400
#endif

#include <atlbase.h>
#include <atlapp.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>
#include <atlctl.h>
#include <atlctrls.h>
#include <atlmisc.h>
#include <atlgdix.h>
#include <atlframe.h>

#include "MemHelper.h"
#include <algorithm>
#include "ComBool.h"
#include "ToolbarHelper.h"

template <class EnumType, class CollType>
HRESULT CreateSTLEnumerator(IUnknown** ppUnk, IUnknown* pUnkForRelease, CollType& collection)
{
	HRESULT hr = E_POINTER;
	if (ppUnk != NULL)
	{
		*ppUnk = NULL;		
		CComObject<EnumType>* pEnum = NULL;
		if (SUCCEEDED(hr = CComObject<EnumType>::CreateInstance(&pEnum)))
		{		
			if (SUCCEEDED(hr = pEnum->Init(pUnkForRelease, collection)))
			{
				hr = pEnum->QueryInterface(ppUnk);
			}
		}
		
		if (FAILED(hr) && pEnum)
			DELETEP(pEnum);
	}
	
	return hr;
	
} // CreateSTLEnumerator

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__8A41C9D2_09C3_422B_B832_438E87D7ED55__INCLUDED)
