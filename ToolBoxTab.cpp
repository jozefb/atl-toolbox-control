// ToolBoxTab.cpp : Implementation of CToolBoxTab
#include "stdafx.h"
#include "IfsToolbox.h"
#include "ToolBoxTab.h"
#include "ToolBoxTabs.h"
#include "ToolBoxItems.h"
#include "ToolBox.h"

#define CE_TBT_CHECK_PTR(ptr, uiResStr, error) if( !(ptr) ) {  return Error(uiResStr, IID_IToolBoxTab, error); }
#define CE_TBT_CHECK_BOOL(b, uiResStr, error) if( !(b) ) {  return Error(uiResStr, IID_IToolBoxTab, error); }
#define CE_TBT_CHECK_HR(hr, uiResStr) CE_CHECK_HR(hr, IID_IToolBoxTab, uiResStr)
#define CE_TBT_CHECK_PTR_ARG(ptr) if( !(ptr) ) {  return Error(IDS_E_INVALID_ARG, IID_IToolBoxTab, E_INVALIDARG); }

/////////////////////////////////////////////////////////////////////////////
// CToolBoxTab

CToolBoxTab::CToolBoxTab() : 
	m_bListView(true),
	m_pTabs(NULL),
	m_nID(-1),
	m_bDeleteable(true),
	m_bHidden(false),
	m_pItems(NULL)
{
}

CToolbox* CToolBoxTab::GetToolbox()
{
	return m_pTabs->GetToolbox();
}

CToolBarCtrl& CToolBoxTab::GetToolBarCtrl()
{
	return m_pTabs->GetToolbox()->GetToolBarCtrl(m_nID);
}

HRESULT CToolBoxTab::GetItem(int nItemID, IToolBoxItem** ppItem)
{
	ATLASSERT(m_pItems);
	return m_pItems->GetItem(nItemID, ppItem);
}

STDMETHODIMP CToolBoxTab::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IToolBoxTab
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CToolBoxTab::FinalConstruct()
{
	CComObject<CToolBoxItems>::CreateInstance(&m_pItems);
	ADDREF(m_pItems);
	m_pItems->put_Parent(this);
		
	return S_OK;
}

STDMETHODIMP CToolBoxTab::FinalRelease()
{
	HRESULT hr=S_OK;

	RELEASE(m_pItems);
	return hr;
}

STDMETHODIMP CToolBoxTab::get_Name(BSTR *pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);

	return m_bstrName.CopyTo(pVal);
}

STDMETHODIMP CToolBoxTab::put_Name(BSTR newVal)
{
	CE_TBT_CHECK_PTR_ARG(newVal);
	if( m_bstrName != newVal )
	{
		if( m_pTabs )
		{
			m_pTabs->GetToolbox()->put_TabName(m_nID, newVal);
		}
		m_bstrName= newVal;
		SetDirty(true);
	}

	return S_OK;
}

STDMETHODIMP CToolBoxTab::get_Guid(BSTR *pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	return m_bstrGuid.CopyTo(pVal);
}

STDMETHODIMP CToolBoxTab::put_Guid(BSTR newVal)
{
	CE_TBT_CHECK_PTR_ARG(newVal);
	if( m_bstrGuid != newVal )
	{
		m_bstrGuid = newVal;
		SetDirty(true);
	}
	
	return S_OK;
}

STDMETHODIMP CToolBoxTab::Delete()
{
	ATLASSERT(m_pTabs);

	return m_pTabs->RemoveTab(m_nID);
}

STDMETHODIMP CToolBoxTab::Activate()
{
	ATLASSERT(m_pTabs);

	return m_pTabs->ActivateTab(m_nID);
}

STDMETHODIMP CToolBoxTab::get_Collection(IToolBoxTabs **pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pTabs);

	return m_pTabs->QueryInterface(__uuidof(IToolBoxTabs), (void**)pVal);
}

STDMETHODIMP CToolBoxTab::get_ListView(VARIANT_BOOL *pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	*pVal=m_bListView;
	return S_OK;
}

STDMETHODIMP CToolBoxTab::put_ListView(VARIANT_BOOL newVal)
{
	if( m_bListView != newVal )
	{
		m_bListView = newVal;
		CToolBarCtrl& tb = GetToolBarCtrl();		
		TBBUTTONINFO tbi = { 0 };
		tbi.cbSize = sizeof(TBBUTTONINFO);
		TBBUTTON btn = { 0 };
		DWORD dwRemove = m_bListView ? TBSTYLE_WRAPABLE : TBSTYLE_LIST;
		DWORD dwAdd = m_bListView ? TBSTYLE_LIST : TBSTYLE_WRAPABLE;
		tb.ModifyStyle(dwRemove, dwAdd, SWP_DRAWFRAME|SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE|SWP_NOZORDER);
		CRect rc;
		if( m_bListView )
		{			
			GetToolbox()->GetClientRect(&rc);
			
			for( int j=0; j<tb.GetButtonCount(); j++ ) 
			{
				tb.GetButton(j, &btn);
				tbi.dwMask = TBIF_SIZE;
				tbi.cx = rc.Width()-2;
				tb.SetButtonInfo(btn.idCommand, &tbi);
			}

			tb.SetRows(tb.GetButtonCount(), TRUE, &rc);
			tb.SetMaxTextRows(1);
		}
		else
		{			
			for( int j=0; j<tb.GetButtonCount(); j++ )
			{				
				tb.GetButton(j, &btn);
				tbi.dwMask = TBIF_SIZE;
				tbi.cx=GetToolbox()->GetItemWidth();
				tb.SetButtonInfo(btn.idCommand, &tbi);
			}

			
			tb.SetRows(1, TRUE, &rc);
			tb.SetMaxTextRows(0);
		}

		SetDirty(true);
		GetToolbox()->Fire_TabListViewChanged(m_nID, m_bstrName, newVal);
	}
	
	return S_OK;
}

void CToolBoxTab::put_Collection(CToolBoxTabs *pTabs)
{
	ATLASSERT(pTabs);
	ATLASSERT(m_nID >=0);
	m_pTabs=pTabs;
	CToolBarCtrl& tb = GetToolBarCtrl();
	HIMAGELIST hImageList = ImageList_LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDB_TAB_BASE), 16, 1, CLR_DEFAULT, IMAGE_BITMAP, LR_CREATEDIBSECTION | LR_DEFAULTSIZE);

	tb.SetImageList(hImageList);
	tb.SetPadding(1, 4);
	// I would like these text styles...
	tb.SetDrawTextFlags(DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_CENTER | DT_BOTTOM, 
		DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
	// and these border colours when the mouse hovers over the toolbutton
	COLORSCHEME color;
	color.dwSize = sizeof(COLORSCHEME);
	color.clrBtnShadow = ::GetSysColor(COLOR_BTNSHADOW);//CLR_DEFAULT 
	color.clrBtnHighlight = ::GetSysColor(COLOR_BTNHIGHLIGHT);//CLR_DEFAULT 
	tb.SetColorScheme(&color);
}

STDMETHODIMP CToolBoxTab::get_ID(int *pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	*pVal=m_nID;
	return S_OK;
}

STDMETHODIMP CToolBoxTab::put_ID(int newVal)
{
	ATLASSERT(newVal>=0);
	m_nID=newVal;
	return S_OK;
}

STDMETHODIMP CToolBoxTab::get_DTE(LPDISPATCH *pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pTabs);
	return m_pTabs->get_DTE(pVal);
}


STDMETHODIMP CToolBoxTab::get_ToolBoxItems(IToolBoxItems **pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pItems);
	return m_pItems->QueryInterface(pVal);
}


STDMETHODIMP CToolBoxTab::get_Deleteable(VARIANT_BOOL *pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	*pVal=m_bDeleteable;
	return S_OK;
}

STDMETHODIMP CToolBoxTab::put_Deleteable(VARIANT_BOOL newVal)
{
	if( m_bDeleteable != newVal )
	{
		m_bDeleteable=newVal;
		SetDirty(true);
	}

	return S_OK;
}

STDMETHODIMP CToolBoxTab::get_Hidden(VARIANT_BOOL *pVal)
{
	CE_TBT_CHECK_PTR_ARG(pVal);
	*pVal=m_bHidden;
	return S_OK;
}

STDMETHODIMP CToolBoxTab::put_Hidden(VARIANT_BOOL newVal)
{
	if( m_bHidden != newVal )
	{
		m_bHidden=newVal;
		SetDirty(true);
		return m_pTabs->ShowTab(m_nID, m_bstrName, !m_bHidden);
	}
	
	return S_OK;
}

CToolBoxItems* CToolBoxTab::GetItems()
{
	ATLASSERT(m_pItems);
	return m_pItems;
}


HRESULT CToolBoxTab::GetSelectedItem(IToolBoxItem **ppItem)
{
	ATLASSERT(m_pItems);
	ATLASSERT(ppItem);
	if( !ppItem )
	{
		return E_INVALIDARG;
	}

	return m_pItems->GetSelectedItem(ppItem);
}
