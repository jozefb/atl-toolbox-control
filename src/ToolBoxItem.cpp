// ToolBoxItem.cpp : Implementation of CToolBoxItem
#include "stdafx.h"
#include "IfsToolbox.h"
#include "ToolBoxItem.h"
#include "ToolBoxItems.h"
#include "ToolBox.h"

/////////////////////////////////////////////////////////////////////////////
// CToolBoxItem

#define CE_TBTI_CHECK_PTR(ptr, uiResStr, error) if( !(ptr) ) {  return Error(uiResStr, IID_IToolBoxItem, error); }
#define CE_TBTI_CHECK_BOOL(b, uiResStr, error) if( !(b) ) {  return Error(uiResStr, IID_IToolBoxItem, error); }
#define CE_TBTI_CHECK_HR(hr, uiResStr) CE_CHECK_HR(hr, IID_IToolBoxItem, uiResStr)
#define CE_TBTI_CHECK_PTR_ARG(ptr) if( !(ptr) ) {  return Error(IDS_E_INVALID_ARG, IID_IToolBoxItem, E_INVALIDARG); }

CToolBoxItem::CToolBoxItem() : 
	m_nID(-1),
	m_eFormat(ToolBoxItemFormatText),
	m_pCollection(NULL),
	m_bDeleteable(true)
{
	
}

STDMETHODIMP CToolBoxItem::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IToolBoxItem
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CToolBoxItem::FinalConstruct()
{
		
	return S_OK;
}

STDMETHODIMP CToolBoxItem::FinalRelease()
{
	HRESULT hr=S_OK;
	
	
	
	return hr;
}

STDMETHODIMP CToolBoxItem::get_Name(BSTR *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);
	return m_bstrName.CopyTo(pVal);
}

STDMETHODIMP CToolBoxItem::put_Name(BSTR newVal)
{
	CE_TBTI_CHECK_PTR_ARG(newVal);
	if( m_bstrName != newVal )
	{
		m_bstrName= newVal;

		CToolBoxTab* pTab = NULL;
		
		if( m_pCollection && (pTab = m_pCollection->GetTab()) )
		{
			CToolBarCtrl& tb = pTab->GetToolBarCtrl();
			CString csName = m_bstrName;
			TBBUTTONINFO tbi={0};
			tbi.cbSize = sizeof(tbi);
			tbi.dwMask = TBIF_TEXT;
			tbi.pszText = (LPTSTR)(LPCTSTR)csName;
			tbi.cchText = csName.GetLength();
			tb.SetButtonInfo(m_nID, &tbi);
		}

		SetDirty(true);
	}

	return S_OK;
}

STDMETHODIMP CToolBoxItem::get_Guid(BSTR *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);
	return m_bstrGuid.CopyTo(pVal);
}

STDMETHODIMP CToolBoxItem::put_Guid(BSTR newVal)
{
	CE_TBTI_CHECK_PTR_ARG(newVal);
	if( m_bstrGuid != newVal )
	{
		m_bstrGuid = newVal;
		SetDirty(true);
	}

	return S_OK;
}

STDMETHODIMP CToolBoxItem::Delete()
{
	ATLASSERT(m_pCollection);
	if( m_bDeleteable )
	{
		return m_pCollection->DeleteItem(m_nID);
	}

	return Error(IDS_E_ITEM_NOT_DELETEABLE, IID_IToolBoxItem, E_FAIL);
}

STDMETHODIMP CToolBoxItem::Select()
{
	ATLASSERT(m_pCollection);
	return m_pCollection->SelectItem(m_nID);
}

STDMETHODIMP CToolBoxItem::get_Collection(IToolBoxItems **pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pCollection);
	return m_pCollection->QueryInterface(__uuidof(IToolBoxItems), (void**)pVal);
}


STDMETHODIMP CToolBoxItem::put_Collection(CToolBoxItems *pVal)
{
	ATLASSERT(pVal);
	CE_TBTI_CHECK_PTR_ARG(pVal);
	m_pCollection=pVal;
	return S_OK;
}

STDMETHODIMP CToolBoxItem::get_ID(int *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);
	*pVal=m_nID;
	return S_OK;
}

STDMETHODIMP CToolBoxItem::put_ID(int newVal)
{
	m_nID=newVal;

	return S_OK;
}

STDMETHODIMP CToolBoxItem::get_DTE(LPDISPATCH *pVal)
{
	ATLASSERT(m_pCollection);
	return m_pCollection->get_DTE(pVal);
}

STDMETHODIMP CToolBoxItem::get_Format(ToolBoxItemFormat *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);
	*pVal=m_eFormat;
	return S_OK;
}

STDMETHODIMP CToolBoxItem::put_Format(ToolBoxItemFormat newVal)
{
	m_eFormat=newVal;
	return S_OK;
}

STDMETHODIMP CToolBoxItem::get_Data(VARIANT *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);
	return ::VariantCopy(pVal, &m_varData);
}

STDMETHODIMP CToolBoxItem::put_Data(VARIANT newVal)
{
	//CE_TBTI_CHECK_PTR_ARG(pVal);
	if( m_varData != newVal )
	{
		return m_varData.Copy(&newVal);
	}
	
	return S_OK;
}

STDMETHODIMP CToolBoxItem::get_Description(BSTR *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);
	return m_bstrDescription.CopyTo(pVal);
}

STDMETHODIMP CToolBoxItem::put_Description(BSTR newVal)
{
	CE_TBTI_CHECK_PTR_ARG(newVal);
	if( m_bstrDescription != newVal )
	{
		m_bstrDescription = newVal;
		SetDirty(true);
	}
	
	return S_OK;
}

STDMETHODIMP CToolBoxItem::get_Hidden(VARIANT_BOOL *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);	
	*pVal=m_bHidden;
	return S_OK;
}

STDMETHODIMP CToolBoxItem::put_Hidden(VARIANT_BOOL newVal)
{
	if( m_bHidden != newVal )
	{
		m_bHidden=newVal;
		SetDirty(true);
		return m_pCollection->HideItem(m_nID, m_bHidden);
	}
	return S_OK;
}

STDMETHODIMP CToolBoxItem::get_Deleteable(VARIANT_BOOL *pVal)
{
	CE_TBTI_CHECK_PTR_ARG(pVal);	
	*pVal=m_bDeleteable;
	return S_OK;
}

STDMETHODIMP CToolBoxItem::put_Deleteable(VARIANT_BOOL newVal)
{
	if( m_bDeleteable != newVal )
	{
		SetDirty(true);
		m_bDeleteable=newVal;
	}

	return S_OK;
}

bool CToolBoxItem::IsSelected()
{
	ATLASSERT(m_pCollection);
	return m_pCollection->IsItemSelected(m_nID);
}

STDMETHODIMP CToolBoxItem::Save(LPSTREAM pStm, BOOL fClearDirty)
{
	HRESULT hr = IPersistStreamInitImpl<CToolBoxItem>::Save(pStm, fClearDirty);
	ATLASSERT(SUCCEEDED(hr));
	return hr;
	
}


STDMETHODIMP CToolBoxItem::Load(LPSTREAM pStm)
{
	HRESULT hr = IPersistStreamInitImpl<CToolBoxItem>::Load(pStm);
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}

