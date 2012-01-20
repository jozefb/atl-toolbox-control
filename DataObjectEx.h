// DataObjectEx.h: interface for the CDataObjectEx class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DATAOBJECTEX_H__5F0A7163_DBE9_44B4_AC63_C68E820BC2A7__INCLUDED_)
#define AFX_DATAOBJECTEX_H__5F0A7163_DBE9_44B4_AC63_C68E820BC2A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <vector>
#include "EnumFormatEtc.h"

HRESULT _CopyStgMedium(STGMEDIUM* pMedDest, STGMEDIUM* pMedSrc, FORMATETC* pFmtSrc);

//template<class T>
class CDataObjectImpl : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDataObjectImpl<CDataObjectImpl> 
{
	typedef struct tagDATAOBJ
	{
		FORMATETC FmtEtc;
		STGMEDIUM StgMed;
		BOOL bRelease;
	} DATAOBJ;

	static _releaseStgMedium(DATAOBJ& obj)
	{
		if( obj.bRelease )
		{
			::ReleaseStgMedium(&obj.StgMed);
		}
	}

	struct _findFormatEtc
	{
		FORMATETC* pFormatEtc;
		_findFormatEtc(FORMATETC* pFormatEtc)
		{
			ATLASSERT(pFormatEtc);
			this->pFormatEtc=pFormatEtc;
		}

		bool operator()(DATAOBJ& obj)
		{
			if((pFormatEtc->tymed    &  obj.FmtEtc.tymed)   &&
				pFormatEtc->cfFormat == obj.FmtEtc.cfFormat &&
				pFormatEtc->dwAspect == obj.FmtEtc.dwAspect)
			{
				return true;
			}

			return false;
		}
	};

public:

	CDataObjectImpl(){}
	virtual ~CDataObjectImpl()
	{
		std::for_each(m_vecDataObj.begin(), m_vecDataObj.end(), _releaseStgMedium);
		m_vecDataObj.clear();
	}

	BEGIN_COM_MAP(CDataObjectImpl)
		COM_INTERFACE_ENTRY(IDataObject)
	END_COM_MAP()

	STDMETHOD(GetDataHere)(FORMATETC* /* pformatetc */, STGMEDIUM* /* pmedium */)
	{
		ATLTRACENOTIMPL(_T("CDataObjectImpl::GetDataHere"));
	}

	STDMETHOD(GetData)(FORMATETC *pformatetcIn, STGMEDIUM *pmedium)
	{
		ATLTRACE(_T("CDataObjectImpl::GetData\n"));
		return IDataObject_GetData(pformatetcIn, pmedium);
	}

	STDMETHOD(IDataObject_GetData)(FORMATETC *pFormatEtc, STGMEDIUM *pMedium)
	{
		ATLTRACE(_T("CDataObjectImpl::IDataObject_GetData\n"));
		ATLASSERT(pMedium);
		ATLASSERT(pFormatEtc);
		itVecDataObj it = std::find_if(m_vecDataObj.begin(), m_vecDataObj.end(), _findFormatEtc(pFormatEtc));
		if( it == m_vecDataObj.end() )
		{
			return DV_E_FORMATETC;
		}
		
		if( pFormatEtc->lindex != -1 )
		{
			return DV_E_LINDEX;
		}
		
		if( ((*it).FmtEtc.tymed & pFormatEtc->tymed) == 0 )
		{
			return DV_E_LINDEX;
		}
		
		return _CopyStgMedium(pMedium, &(*it).StgMed, &(*it).FmtEtc);
	}

	STDMETHOD(SetData)(FORMATETC* pFormatEtc, STGMEDIUM* pmedium, BOOL fRelease)
	{
		ATLTRACE(_T("CDataObjectImpl::SetData\n"));
		ATLASSERT(pFormatEtc);
		ATLASSERT(pmedium);
		if( pFormatEtc == NULL ) return E_POINTER;
		if( pmedium == NULL ) return E_POINTER;
		DATAOBJ obj;
		obj.FmtEtc = *pFormatEtc;
		obj.StgMed = *pmedium;
		obj.bRelease = fRelease;
/*
		if( !fRelease ) 
		{
			_CopyStgMedium(&obj.StgMed, pmedium, pFormatEtc);
		}
*/
		
		itVecDataObj it = std::find_if(m_vecDataObj.begin(), m_vecDataObj.end(), _findFormatEtc(pFormatEtc));
		if( it == m_vecDataObj.end() ) 
		{
			m_vecDataObj.push_back(obj);
		}
		else
		{
			m_vecDataObj.insert(it, obj);
		}
		
		return S_OK;
	}
	
	STDMETHOD(EnumFormatEtc)(DWORD dwDirection, IEnumFORMATETC **ppEnumFormatEtc)
	{
		ATLTRACE(_T("CDataObjectImpl::EnumFormatEtc\n"));
		ATLASSERT(ppEnumFormatEtc);
		if( ppEnumFormatEtc == NULL ) 
		{
			return E_POINTER;
		}
		
		*ppEnumFormatEtc = NULL;
		if( dwDirection != DATADIR_GET ) 
		{
			return E_NOTIMPL;
		}
		CComObject<CEnumFORMATETC>* pEnum = NULL;
		if( FAILED( CComObject<CEnumFORMATETC>::CreateInstance(&pEnum) ) )
		{
			return E_FAIL;
		}
		
		for( int i = 0; i < m_vecDataObj.size(); i++ ) 
		{
			pEnum->Add(m_vecDataObj[i].FmtEtc);
		}
		
		return pEnum->QueryInterface(ppEnumFormatEtc);
	}
	
	STDMETHOD(QueryGetData)(FORMATETC* pFormatEtc)
	{
		ATLTRACE(_T("CDataObjectImpl::QueryGetData\n"));
		ATLASSERT(pFormatEtc);
		itVecDataObj it = std::find_if(m_vecDataObj.begin(), m_vecDataObj.end(), _findFormatEtc(pFormatEtc));
		if( it == m_vecDataObj.end() ) 
		{
			return DV_E_FORMATETC;
		}
		
		// BUG: Only supports one media pr format!
		if( (*it).FmtEtc.lindex != -1 )
		{
			return DV_E_LINDEX;
		}
		
		if( (*it).FmtEtc.tymed != pFormatEtc->tymed )
		{
			return DV_E_TYMED;
		}
		
		return S_OK;
	}
	
	STDMETHOD(GetCanonicalFormatEtc)(FORMATETC* pformatectIn, FORMATETC* /* pformatetcOut */)
	{
		ATLTRACE(_T("CDataObjectImpl::GetCanonicalFormatEtc\n"));
		pformatectIn->ptd = NULL;
		ATLTRACENOTIMPL(_T("CDataObjectEx::GetCanonicalFormatEtc"));
	}

	HRESULT SetGlobalData(CLIPFORMAT cf, LPCVOID pData, DWORD dwSize)
	{
		ATLTRACE(_T("CDataObjectImpl::SetGlobalData\n"));
		FORMATETC fmtetc = { cf, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		STGMEDIUM stgmed = { TYMED_HGLOBAL, { 0 }, 0 };
		stgmed.hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, (SIZE_T) dwSize);
		if( stgmed.hGlobal == NULL ) return E_OUTOFMEMORY;
		memcpy(GlobalLock(stgmed.hGlobal), pData, (size_t) dwSize);
		GlobalUnlock(stgmed.hGlobal);
		return SetData(&fmtetc, &stgmed, TRUE);
	}

	HRESULT SetIStreamData(CLIPFORMAT cf, IStream* pStream, BOOL fRelease)
	{
		ATLTRACE(_T("CDataObjectImpl::SetIStreamData\n"));
		FORMATETC fmtetc = { cf, 0, DVASPECT_CONTENT, -1, TYMED_ISTREAM };
		STGMEDIUM stgmed = {0};
		stgmed.tymed = TYMED_ISTREAM;
		stgmed.pstm = pStream;

		return SetData(&fmtetc, &stgmed, fRelease);
	}

	HRESULT SetTextData(LPCSTR pstrData)
	{
		return SetGlobalData(CF_TEXT, pstrData, _tcsclen(pstrData) + 1);
	}
	HRESULT SetUnicodeTextData(CComBSTR bstrData)
	{
		return SetGlobalData(CF_UNICODETEXT, bstrData, bstrData.Length());
	}

	CComPtr<IDataAdviseHolder> m_spDataAdviseHolder;
	
protected:
	std::vector<DATAOBJ> m_vecDataObj;
	typedef std::vector<DATAOBJ>::iterator itVecDataObj;
	

};

#endif // !defined(AFX_DATAOBJECTEX_H__5F0A7163_DBE9_44B4_AC63_C68E820BC2A7__INCLUDED_)
