// EnumFormatEtc.cpp: implementation of the CEnumFORMATETC class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "EnumFormatEtc.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

HRESULT DeepCopyFormatEtc(FORMATETC *dest, FORMATETC *source)
{
	// copy the source FORMATETC into dest
	*dest = *source;
	
	if(source->ptd&& dest->ptd)
	{
		// allocate memory for the DVTARGETDEVICE if necessary
		dest->ptd = (DVTARGETDEVICE*)COMEMALLOC(sizeof(DVTARGETDEVICE));
		if( !dest->ptd )
		{
			return E_POINTER;
		}
		// copy the contents of the source DVTARGETDEVICE into dest->ptd
		*(dest->ptd) = *(source->ptd);
	}
	return S_OK;
    
}

CEnumFORMATETC::~CEnumFORMATETC()
{
		
}

STDMETHODIMP CEnumFORMATETC::Reset()
{
    m_itCurent = m_vecFormats.begin();
    return S_OK;
}

STDMETHODIMP CEnumFORMATETC::Skip(ULONG celt)
{
	HRESULT hr = S_OK;
	while (celt--)
	{
		if (m_itCurent != m_vecFormats.end())
			m_itCurent++;
		else
		{
			hr = S_FALSE;
			break;
		}
	}
	return hr;
}

STDMETHODIMP CEnumFORMATETC::Clone(IEnumFORMATETC **ppEnumFormatEtc)
{
	HRESULT hRes = E_POINTER;
	if (ppEnumFormatEtc != NULL)
	{
		*ppEnumFormatEtc = NULL;
		CComObject<CEnumFORMATETC>* p;
		hRes = CComObject<CEnumFORMATETC>::CreateInstance(&p);
		if (SUCCEEDED(hRes))
		{
			//hRes = p->Init(m_spUnk, *m_pcollection);
			if (SUCCEEDED(hRes))
			{
				p->m_itCurent = m_itCurent;
				p->m_vecFormats = m_vecFormats;
				hRes = p->_InternalQueryInterface(__uuidof(IEnumFORMATETC), (void**)ppEnumFormatEtc);
			}
			if (FAILED(hRes))
				delete p;
		}
	}
	return hRes;
}

STDMETHODIMP CEnumFORMATETC::Next(ULONG celt, FORMATETC *lpFormatEtc, ULONG *pceltFetched)
{
    if( pceltFetched != NULL) 
	{
		*pceltFetched = 0;
	}

	if( !lpFormatEtc || celt != 1 )
	{
		return E_INVALIDARG;
	}

	if( celt <= 0 || m_itCurent == m_vecFormats.end() ) 
	{
		return S_FALSE;
	}

	ULONG nCount = 0;
	while( m_itCurent != m_vecFormats.end() && celt > 0 ) 
	{
		*lpFormatEtc++ = *(m_itCurent)++;
		--celt;
		++nCount;
	}
	if( pceltFetched != NULL ) 
	{
		*pceltFetched = nCount;
	}
	
	return celt == 0 ? S_OK : S_FALSE;
}

void CEnumFORMATETC::Add(FORMATETC &fmtc)
{
	m_vecFormats.push_back(fmtc);
}
