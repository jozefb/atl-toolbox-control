// EnumFormatEtc.h: interface for the CEnumFORMATETC class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ENUMFORMATETC_H__F1692451_01AE_415F_96AE_ACFC857DB612__INCLUDED_)
#define AFX_ENUMFORMATETC_H__F1692451_01AE_415F_96AE_ACFC857DB612__INCLUDED_

#include <vector>

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class ATL_NO_VTABLE CEnumFORMATETC : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public IEnumFORMATETC  
{
public:
	CEnumFORMATETC()
	{

	}
	
	virtual ~CEnumFORMATETC();

	BEGIN_COM_MAP(CEnumFORMATETC)
		COM_INTERFACE_ENTRY(IEnumFORMATETC)
	END_COM_MAP()

	STDMETHOD(Reset)(void);
	STDMETHOD(Skip)(ULONG celt);
	STDMETHOD(Clone)(IEnumFORMATETC **ppEnumFormatEtc);
	STDMETHOD(Next)(ULONG celt, FORMATETC *pFormatEtc, ULONG *pceltFetched);
	

	void Add(FORMATETC& fmtc);

protected:
	std::vector<FORMATETC> m_vecFormats;
	typedef std::vector<FORMATETC>::iterator itVecFormats;
	itVecFormats m_itCurent;
	
};

#endif // !defined(AFX_ENUMFORMATETC_H__F1692451_01AE_415F_96AE_ACFC857DB612__INCLUDED_)
