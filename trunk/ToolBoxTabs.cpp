// ToolBoxTabs.cpp : Implementation of CToolBoxTabs
#include "stdafx.h"
#include "IfsToolbox.h"
#include "ToolBoxTabs.h"
#include "Toolbox.h"


#define CE_TBTS_CHECK_PTR(ptr, uiResStr, error) if( !(ptr) ) {  return Error(uiResStr, __uuidof(IToolBoxTabs), error); }
#define CE_TBTS_CHECK_BOOL(b, uiResStr, error) if( !(b) ) {  return Error(uiResStr, __uuidof(IToolBoxTabs), error); }
#define CE_TBTS_CHECK_HR(hr, uiResStr) CE_CHECK_HR(hr, __uuidof(IToolBoxTabs), uiResStr)
#define CE_TBTS_CHECK_PTR_ARG(ptr) if( !(ptr) ) {  return Error(IDS_E_INVALID_ARG, __uuidof(IToolBoxTabs), E_INVALIDARG); }
#define COM_ERR(resstr, err) Error(resstr, __uuidof(IToolBoxTabs), err);


/////////////////////////////////////////////////////////////////////////////
// CToolBoxTabs

STDMETHODIMP CToolBoxTabs::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IToolBoxTabs
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CToolBoxTabs::FinalConstruct()
{
		
	return S_OK;
}

STDMETHODIMP CToolBoxTabs::FinalRelease()
{
	HRESULT hr=S_OK;

	std::for_each(m_vecTabs.begin(), m_vecTabs.end(), _ReleaseTab);
	
	m_vecTabs.clear();
	
	return hr;
}

STDMETHODIMP CToolBoxTabs::get_Parent(IToolbox* *pVal)
{
	CE_TBTS_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pToolbox);
	return m_pToolbox->QueryInterface(__uuidof(IToolbox), (void**)pVal);
}

STDMETHODIMP CToolBoxTabs::put_Parent(CToolbox *pVal)
{
	CE_TBTS_CHECK_PTR_ARG(pVal);
	m_pToolbox = pVal;
	return S_OK;
}

STDMETHODIMP CToolBoxTabs::get_DTE(LPDISPATCH *pVal)
{
	CE_TBTS_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pToolbox);
	return m_pToolbox->get_DTE(pVal);
}

STDMETHODIMP CToolBoxTabs::get__NewEnum(LPUNKNOWN *pVal)
{	
	CE_TBTS_CHECK_PTR_ARG(pVal);
	
	typedef CComEnumOnSTL<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT,
		_CopyIDispatch, VectorTabs > CComEnumUnknownOnVector;
	
	return CreateSTLEnumerator<CComEnumUnknownOnVector>(pVal, (IToolBoxTabs*)this, m_vecTabs);
}

STDMETHODIMP CToolBoxTabs::get_Item(VARIANT index, IToolBoxTab **pVal)
{
	CE_TBTS_CHECK_PTR_ARG(pVal);
	
	if( !(index.vt==VT_BSTR||index.vt==VT_I4) ) 
	{  
		return COM_ERR(IDS_E_INVALID_ARG, E_INVALIDARG);
	}
	
	if( index.vt==VT_I4 )
	{
		if( !(index.lVal>=0 && index.lVal<m_vecTabs.size()) ) 
		{  
			return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
		}
		
		return m_vecTabs[index.lVal].pTab->QueryInterface(__uuidof(IToolBoxTab), (void**)pVal);
	}
	
	if( index.vt==VT_BSTR )
	{
		CE_TBTS_CHECK_PTR_ARG(index.bstrVal);
		vecIter it = std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findTabByName(index.bstrVal));
		if( it != m_vecTabs.end() )
		{
			return (*it).pTab->QueryInterface(__uuidof(IToolBoxTab), (void**)pVal);
		}
	}

	return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
}

STDMETHODIMP CToolBoxTabs::get_Count(long *pVal)
{
	CE_TBTS_CHECK_PTR_ARG(pVal);
	*pVal=m_vecTabs.size();
	return S_OK;
}

STDMETHODIMP CToolBoxTabs::Add(BSTR bstrName, IToolBoxTab **ppToolBoxTab)
{
	//CE_TBTS_CHECK_PTR_ARG(ppToolBoxTab);
	CE_TBTS_CHECK_PTR_ARG(bstrName);
	if( TabWithNameExist(bstrName) )
	{
		CString csFmt;
		csFmt.LoadString(IDS_E_TAB_WITHNAME_EXIST);
		CString csName=bstrName;
		CString csErr;
		csErr.Format(csFmt, csName);
		return COM_ERR(csErr, E_INVALIDARG);
	}

	CComObject<CToolBoxTab>* pTab=NULL;
	CComObject<CToolBoxTab>::CreateInstance(&pTab);
	ADDREF(pTab);

	int nID = m_pToolbox->InsertTab(-1, bstrName);
	pTab->put_ID(nID);
	pTab->put_Collection(this);
	pTab->put_Name(bstrName);
	pTab->SetDirty(false);
	m_vecTabs.push_back(_tTabHolder(pTab));

	if( ppToolBoxTab )
	{
		pTab->QueryInterface(ppToolBoxTab);
	}
	m_pToolbox->Fire_TabAdded(nID, bstrName);

	return S_OK;
}

bool CToolBoxTabs::TabWithNameExist(BSTR bstrName)
{
	return std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findTabByName(bstrName)) != m_vecTabs.end();;
}

HRESULT CToolBoxTabs::RemoveTab(int nTabID)
{
	ATLASSERT(m_pToolbox);
	CE_TBTS_CHECK_BOOL(nTabID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	vecIter it = std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findTabByID(nTabID));
	if( it != m_vecTabs.end() )
	{
		CComObject<CToolBoxTab>* pTab=(*it).pTab;
		CComBSTR bstrName;
		pTab->get_Name(&bstrName);
		m_vecTabs.erase(it);
		RELEASE(pTab);
		ATLVERIFY(m_pToolbox->DeleteTab(nTabID));
		m_pToolbox->Fire_TabRemoved(nTabID, bstrName);
		return S_OK;
	}

	return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
}

HRESULT CToolBoxTabs::ActivateTab(int nTabID)
{
	ATLASSERT(m_pToolbox);
	CE_TBTS_CHECK_BOOL(nTabID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);

	vecIter it = std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findTabByID(nTabID));
	if( it != m_vecTabs.end() )
	{
		CComBSTR bstrName;
		(*it).pTab->get_Name(&bstrName);
		m_pToolbox->MaximizeTab(nTabID);
		m_pToolbox->Fire_TabActivated(nTabID, bstrName);
		m_pToolbox->SetFocus();
		return S_OK;
	}

	return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
}

CToolbox* CToolBoxTabs::GetToolbox()
{
	return m_pToolbox;
}


HRESULT CToolBoxTabs::ShowTab(int nTabID, BSTR Name, bool Show)
{
	ATLASSERT(m_pToolbox);
	CE_TBTS_CHECK_BOOL(nTabID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	
	vecIter it = std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findTabByID(nTabID));
	if( it != m_vecTabs.end() )
	{
		m_pToolbox->ShowTab(nTabID, Show);
		m_pToolbox->Fire_TabHidden(nTabID, Name, Show);
		return S_OK;
	}
	
	return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
}

HRESULT CToolBoxTabs::GetTabByID(int nTabID, IToolBoxTab** pTab)
{
	ATLASSERT(m_pToolbox);
	CE_TBTS_CHECK_BOOL(nTabID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	
	vecIter it = std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findTabByID(nTabID));
	if( it != m_vecTabs.end() )
	{
		return (*it).pTab->QueryInterface(__uuidof(IToolBoxTab), (void**)pTab);
	}
	
	return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
}

CComObject<CToolBoxTab>* CToolBoxTabs::GetTabByID(int nTabID)
{
	ATLASSERT(m_pToolbox);	
	vecIter it = std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findTabByID(nTabID));
	if( it != m_vecTabs.end() )
	{
		return (*it).pTab;
	}
	
	return NULL;
}

HRESULT CToolBoxTabs::GetItemByID(int nItemID, IToolBoxItem** pItem)
{
	ATLASSERT(m_pToolbox);
	CE_TBTS_CHECK_PTR_ARG(pItem);
	CE_TBTS_CHECK_BOOL(nItemID>=MIN_TBITEM_ID, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	
	for(vecIter it = m_vecTabs.begin(); it != m_vecTabs.end(); it++ )
	{
		(*it).pTab->GetItem(nItemID, pItem);
		if( *pItem != NULL )
		{
			return S_OK;
		}
	}
	
	return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
}

HRESULT CToolBoxTabs::GetItemByID(int nTabID, int nItemID, IToolBoxItem** pItem)
{
	ATLASSERT(m_pToolbox);
	CE_TBTS_CHECK_PTR_ARG(pItem);
	CE_TBTS_CHECK_BOOL(nItemID>=MIN_TBITEM_ID, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	CE_TBTS_CHECK_BOOL(nTabID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);

	CComObject<CToolBoxTab>* pTab;
	pTab = GetTabByID(nTabID);
	if( pTab )
	{
		return pTab->GetItem(nItemID, pItem);
	}
	
	return COM_ERR(IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
}

void CToolBoxTabs::ShowAllTabs()
{
	std::for_each(m_vecTabs.begin(), m_vecTabs.end(), _ShowTab);
}

bool CToolBoxTabs::ContainsHiddenItems()
{
	return std::find_if(m_vecTabs.begin(), m_vecTabs.end(), findHiddenTab()) != m_vecTabs.end();
}
