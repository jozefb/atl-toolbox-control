// Toolbox.cpp : Implementation of CToolbox

#include "stdafx.h"
#include "IfsToolbox.h"
#include "Toolbox.h"
#include "ToolBoxItems.h"


#define CE_TBX_CHECK_PTR(ptr, uiResStr, error) if( !(ptr) ) {  return Error(uiResStr, IID_IToolbox, error); }
#define CE_TBX_CHECK_BOOL(b, uiResStr, error) if( !(b) ) {  return Error(uiResStr, IID_IToolbox, error); }
#define CE_TBX_CHECK_HR(hr, uiResStr) CE_CHECK_HR(hr, IID_IToolbox, uiResStr)
#define CE_TBX_CHECK_PTR_ARG(ptr) if( !(ptr) ) {  return Error(IDS_E_INVALID_ARG, IID_IToolbox, E_INVALIDARG); }

static FONTDESC g_FontDesc = { sizeof(FONTDESC), 
OLESTR("Times New Roman"), 
FONTSIZE( 9 ),
FW_NORMAL, 
ANSI_CHARSET, 
TRUE, 
FALSE, 
FALSE 
};

UINT CF_TOOLBOX_ITEM = ::RegisterClipboardFormat("ToolboxItem");

/////////////////////////////////////////////////////////////////////////////
// CToolbox

CToolbox::CToolbox() : 
	m_pTabs(NULL),
	m_wndEdit(_T("EDIT"), this, 1),
	m_bShowPopupMenu(true),
	m_cmdManager(*this),
	m_bPMICustomize(true)
{
	CComControlBase::m_bWindowOnly = TRUE;
}

LRESULT CToolbox::OnRebarLayoutChanged(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
	m_cmdManager.StopCmdProcessing();
	return 0;
}

LRESULT CToolbox::OnRelayToolTip(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
	LPTOOLTIPTEXT lpTTText = (LPTOOLTIPTEXT)pnmh;
	CComPtr<IToolBoxItem> pItem;
	if( SUCCEEDED(m_pTabs->GetItemByID(idCtrl, &pItem)) )
	{
		ATLVERIFY(pItem);
		CComBSTR bstrDescr;
		pItem->get_Description(&bstrDescr);
		CString csDscr=bstrDescr;

		_tcsncpy( lpTTText->szText,csDscr, sizeof(lpTTText->szText) );
	}
	
	return 0;
}

LRESULT CToolbox::OnCharEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	TCHAR chCharCode = (TCHAR) wParam; 
	long lKeyData = lParam;

	bHandled=FALSE;
	return 0;
}

LRESULT CToolbox::OnCtlColorEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{	
	return reinterpret_cast< LRESULT >( ::GetSysColorBrush(COLOR_WINDOW) );
}

LRESULT CToolbox::OnGetDlgCode(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	return DLGC_HASSETSEL|DLGC_WANTALLKEYS|DLGC_WANTCHARS|DLGC_WANTMESSAGE|DLGC_WANTARROWS|DLGC_WANTTAB;
}

LRESULT CToolbox::OnGetDlgCodeEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	return DLGC_HASSETSEL|DLGC_WANTALLKEYS|DLGC_WANTCHARS|DLGC_WANTMESSAGE|DLGC_WANTARROWS|DLGC_WANTTAB;
}

BOOL CToolbox::PreTranslateAccelerator(LPMSG pMsg, HRESULT& hRet)
{
	hRet = S_OK;
	if ((pMsg->message < WM_KEYFIRST || pMsg->message > WM_KEYLAST) &&
		(pMsg->message < WM_MOUSEFIRST || pMsg->message > WM_MOUSELAST))
		return FALSE;
	
	if( m_accel.TranslateAccelerator(m_hWnd, pMsg) )
	{
		return TRUE;
	}
	
	return FALSE;
}

STDMETHODIMP CToolbox::FinalConstruct()
{
	HRESULT hr=CComObject<CToolBoxTabs>::CreateInstance(&m_pTabs);
	ADDREF(m_pTabs);
	m_pTabs->put_Parent(this);

	return hr;
}

STDMETHODIMP CToolbox::FinalRelease()
{
	HRESULT hr=S_OK;

	RELEASE(m_pTabs);
	m_pFont.Release();
	
	return hr;
}

STDMETHODIMP CToolbox::OnAmbientPropertyChange(DISPID dispid)
{
	if(dispid == DISPID_AMBIENT_USERMODE)
	{
		BOOL bUserMode = FALSE;
		GetAmbientUserMode(bUserMode);
		EnableWindow(bUserMode);				
	}
		
		
	return S_OK;
}

STDMETHODIMP CToolbox::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IToolbox,
	};
	for (int i=0; i<sizeof(arr)/sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i], riid))
			return S_OK;
	}
	return S_FALSE;
}

HRESULT CToolbox::OnDraw(ATL_DRAWINFO& di)
{
	return S_OK;
}

HRESULT CToolbox::FireOnChanged(DISPID dispID)
{
	CComControl<CToolbox>::FireOnChanged(dispID);
	switch(dispID) {
	case DISPID_FONT:
		_setupFonts();
		break;
	case DISPID_ENABLED:
		{
			BOOL bUserMode = FALSE;
			GetAmbientUserMode(bUserMode);
			if( IsWindow() )
			{
				EnableWindow( bUserMode ? m_bEnabled : FALSE);
			}
		}
		break;
	default:
		break;
	}
	
	return S_OK;
}

LRESULT CToolbox::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	ModifyStyle(0, WS_CLIPCHILDREN|WS_CLIPSIBLINGS);
	CRect rc;
	GetClientRect(&rc);
	ATLVERIFY(m_CmdBar.Create(m_hWnd, rcDefault, NULL, ATL_SIMPLE_CMDBAR_PANE_STYLE)!=NULL);
	ATLVERIFY(m_CmdBar.SetImageSize(16, 16));	
	
	CBitmap bmp;
	ATLVERIFY(bmp.LoadBitmap(IDB_DELETE));
	ATLVERIFY(m_CmdBar.AddBitmap(bmp, ID_DELETE_TAB));
	ATLVERIFY(bmp.DeleteObject());
	ATLVERIFY(bmp.LoadBitmap(IDB_MOVE_UP));
	ATLVERIFY(m_CmdBar.AddBitmap(bmp, ID_MOVE_UP));
	ATLVERIFY(bmp.DeleteObject());
	ATLVERIFY(bmp.LoadBitmap(IDB_MOVE_DOWN));
	ATLVERIFY(m_CmdBar.AddBitmap(bmp, ID_MOVE_DOWN));
	ATLVERIFY(bmp.DeleteObject());
	ATLVERIFY(bmp.LoadBitmap(IDB_DELETE));
	ATLVERIFY(m_CmdBar.AddBitmap(bmp, ID_DELETE_TABITEM));

	ATLVERIFY(bmp.DeleteObject());
	ATLVERIFY(bmp.LoadBitmap(IDB_CUT));
	ATLVERIFY(m_CmdBar.AddBitmap(bmp, ID_EDIT_CUT));

	ATLVERIFY(bmp.DeleteObject());
	ATLVERIFY(bmp.LoadBitmap(IDB_COPY));
	ATLVERIFY(m_CmdBar.AddBitmap(bmp, ID_EDIT_COPY));

	ATLVERIFY(bmp.DeleteObject());
	ATLVERIFY(bmp.LoadBitmap(IDB_PASTE));
	ATLVERIFY(m_CmdBar.AddBitmap(bmp, ID_EDIT_PASTE));

	m_CmdBar.Prepare();

	m_wndRebar.Create(m_hWnd, rc, NULL, WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN | WS_CLIPSIBLINGS, WS_EX_STATICEDGE);
	m_wndEdit.Create(m_hWnd, rcDefault, "", WS_CHILD|ES_NOHIDESEL|ES_LEFT|ES_AUTOHSCROLL, WS_EX_CLIENTEDGE);
	ATLASSERT(m_pFont);

	_setupFonts();

	ATLVERIFY(m_accel.LoadAccelerators(IDR_MAIN)!=NULL);
	
	BOOL bUserMode = FALSE;
	GetAmbientUserMode(bUserMode);
	if( IsWindow() )
	{
		EnableWindow( bUserMode ? m_bEnabled : FALSE);
	}

	return 0;
}

STDMETHODIMP CToolbox::Save(LPSTREAM pStm, BOOL fClearDirty)
{
	HRESULT hr = IPersistStreamInitImpl<CToolbox>::Save(pStm, fClearDirty);
	ATLASSERT(SUCCEEDED(hr));
	return hr;
	
}


STDMETHODIMP CToolbox::Load(LPSTREAM pStm)
{
	HRESULT hr = IPersistStreamInitImpl<CToolbox>::Load(pStm);
	ATLASSERT(SUCCEEDED(hr));
	return hr;
}

LRESULT CToolbox::OnContextMenu(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	m_cmdManager.StopCmdProcessing();
	if( !m_bShowPopupMenu )
	{
		return 0;
	}

	POINT pt = {  GET_X_LPARAM(lParam),  GET_Y_LPARAM(lParam) };
	m_wndRebar.ScreenToClient(&pt);
	
	RBHITTESTINFO rbhi;
	rbhi.pt=pt;
	int nTabIndex=m_wndRebar.HitTest(&rbhi);
	
	REBARBANDINFO rbi={0};
	rbi.cbSize = sizeof(rbi);
	rbi.fMask=RBBIM_ID;
	int nPopupTabID = m_wndRebar.GetBandInfo(nTabIndex, &rbi) ? rbi.wID : -1;
		
	m_wndRebar.ClientToScreen(&pt);
	CMenu menu;
	ATLVERIFY(menu.LoadMenu(rbhi.flags != RBHT_CLIENT ? IDR_TBX_TAB_POPUP : IDR_TBX_TABITEM_POPUP));
	CMenuHandle popup=menu.GetSubMenu(0);

	if( !m_bPMICustomize )
	{
		popup.RemoveMenu(ID_CUSTOMIZE_TBX, MF_BYCOMMAND);
	}

	long lTabCount=0;
	m_pTabs->get_Count(&lTabCount);

	CComObject<CToolBoxTab>* pTab = m_pTabs->GetTabByID(nPopupTabID);
	ADDREF(pTab);
	CComBool listView;
	UINT nFlags=MF_BYCOMMAND;
	if( pTab )
	{	
		pTab->get_ListView(&listView);		
	}

	menu.CheckMenuItem(ID_LIST_VIEW, MF_BYCOMMAND | (listView ? MF_CHECKED:MF_UNCHECKED));
	menu.CheckMenuItem(ID_SHOW_ALL_TABS, MF_BYCOMMAND | m_pTabs->ContainsHiddenItems() ? MF_UNCHECKED : MF_CHECKED);
	
	if( rbhi.flags != RBHT_CLIENT )
	{
		m_cmdManager.SetTabAndItemID(nPopupTabID, -1);
		m_wndRebar.SetHighlightedBandID(nPopupTabID);
		_setupTabPopupMenu(popup, nTabIndex, pTab);
	}
	else
	{
		CToolBarCtrl& tb = m_wndRebar.GetToolBarCtrl(m_wndRebar.IdToIndex(nPopupTabID));
		POINT ptTmp=pt;
		tb.ScreenToClient(&ptTmp);
		int nItemIndex = tb.HitTest(&ptTmp);
		_setupItemPopupMenu(popup, pTab, nItemIndex, tb);
	}

	RELEASE(pTab);

	m_CmdBar.DoTrackPopupMenu(popup, TPM_LEFTALIGN | TPM_LEFTBUTTON, pt.x, pt.y);
	m_wndRebar.SetHighlightedBandID(-1);

	return 0;
}

void CToolbox::_setupTabPopupMenu(CMenuHandle &menu, int nTabIdx, IToolBoxTab* pTab)
{
	int nCount = m_wndRebar.GetBandCount();
		
	UINT flags = MF_BYCOMMAND;
	CComBool canBeDeleted(false);
	if( pTab )
	{
		pTab->get_Deleteable(&canBeDeleted);
		flags |= canBeDeleted ? MF_ENABLED : MF_GRAYED;
		menu.EnableMenuItem(ID_DELETE_TAB, flags);

		if( nTabIdx == nCount-1 || nCount == 1)
		{
			menu.EnableMenuItem(ID_MOVE_DOWN, MF_BYCOMMAND|MF_GRAYED);
		} 
		if( nTabIdx == 0 || nCount == 1)
		{
			menu.EnableMenuItem(ID_MOVE_UP, MF_BYCOMMAND|MF_GRAYED);
		}
	}
	else
	{
		flags |= MF_GRAYED;
		menu.EnableMenuItem(ID_RENAME_TAB, flags);
		menu.EnableMenuItem(ID_DELETE_TAB, flags);
		menu.EnableMenuItem(ID_HIDE_TAB, flags);
		menu.EnableMenuItem(ID_LIST_VIEW, flags);
		menu.EnableMenuItem(ID_MOVE_DOWN, flags);
		menu.EnableMenuItem(ID_MOVE_UP, flags);
		menu.EnableMenuItem(ID_EDIT_COPY, flags);
		menu.EnableMenuItem(ID_EDIT_PASTE, flags);
		menu.EnableMenuItem(ID_EDIT_CUT, flags);
		menu.EnableMenuItem(ID_SHOW_ALL_TABS, MF_BYCOMMAND| nCount ? MF_ENABLED:MF_GRAYED);
	}
}

LRESULT CToolbox::OnCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	int nCode= HIWORD(wParam);
	int wID = LOWORD(wParam); 
	HWND hWndCtrl = (HWND) lParam;
	
	if( hWndCtrl == NULL )
	{
		// from popup menu or accelerator
		// 
		if( m_cmdManager.GetTabID() == -1 && 
			m_cmdManager.GetItemID() == -1 )
		{
			// it is command from accelerator
			int nTabID=m_wndRebar.GetMaximizedBandID();
			int nItemID=-1;
			CComObject<CToolBoxTab>* pTab = m_pTabs->GetTabByID(nTabID);
			ADDREF(pTab);
			if( pTab )
			{
				CComPtr<IToolBoxItem> pItem;
				pTab->GetSelectedItem(&pItem);
				if( pItem )
				{
					pItem->get_ID(&nItemID);
				}
				
				RELEASE(pTab);
			}
			
			m_cmdManager.SetTabAndItemID(nTabID, nItemID);
		}
		
		return m_cmdManager.ProcessCmd(wID);
	}

	if( hWndCtrl == m_wndRebar )
	{
		// command from toolbar button
		m_cmdManager.StopCmdProcessing();
		CComPtr<IToolBoxItem> pItem;
		m_pTabs->GetItemByID(wID, &pItem);
		if( pItem )
		{
			Fire_ItemSelected(wID, pItem);
		}
	}

	return 0;	
}

inline void CToolbox::_setupItemPopupMenu(CMenuHandle &menu, CToolBoxTab* pTab, int nItemIndex, CToolBarCtrl& tb)
{
	ATLASSERT(menu && pTab && tb);
	int nItemID=-1;
	if( nItemIndex >= 0 )
	{
		TBBUTTON button={0};
		ATLVERIFY(tb.GetButton(nItemIndex, &button));
		nItemID = button.idCommand;
		CComPtr<IToolBoxItem> pItem;
		if( SUCCEEDED(pTab->GetItems()->GetItem(nItemID, &pItem)) )
		{
			ATLVERIFY(SUCCEEDED(pItem->Select()));
			int nID=-1;
			pItem->get_ID(&nID);
			CComBool bDeletable;
			pItem->get_Deleteable(&bDeletable);
			if( bDeletable )
			{
				menu.EnableMenuItem(ID_DELETE_TABITEM, MF_BYCOMMAND|MF_ENABLED);
				menu.EnableMenuItem(ID_EDIT_CUT, MF_BYCOMMAND|MF_ENABLED);
			}
			int nBtnCount = tb.GetButtonCount();
			if( nItemIndex == nBtnCount-1 || nBtnCount == 1 || nID == POINTER_ID)
			{
				menu.EnableMenuItem(ID_MOVE_DOWN, MF_BYCOMMAND|MF_GRAYED);
			}
			if( nItemIndex == 0  || nBtnCount == 1 || nID == POINTER_ID)
			{
				menu.EnableMenuItem(ID_MOVE_UP, MF_BYCOMMAND|MF_GRAYED);
			}
			
			if( nID == POINTER_ID)
			{
				menu.EnableMenuItem(ID_EDIT_COPY, MF_BYCOMMAND|MF_GRAYED);
				menu.EnableMenuItem(ID_EDIT_CUT, MF_BYCOMMAND|MF_GRAYED);
			}
		}
	}
	else
	{
		menu.EnableMenuItem(ID_MOVE_DOWN, MF_BYCOMMAND|MF_GRAYED);
		menu.EnableMenuItem(ID_MOVE_UP, MF_BYCOMMAND|MF_GRAYED);
		menu.EnableMenuItem(ID_EDIT_COPY, MF_BYCOMMAND|MF_GRAYED);
		menu.EnableMenuItem(ID_EDIT_PASTE, MF_BYCOMMAND|MF_GRAYED);
		menu.EnableMenuItem(ID_EDIT_CUT, MF_BYCOMMAND|MF_GRAYED);
		menu.EnableMenuItem(ID_DELETE_TABITEM, MF_BYCOMMAND|MF_GRAYED);
		menu.EnableMenuItem(ID_RENAME_TABITEM, MF_BYCOMMAND|MF_GRAYED);
	}

	int nTabID=-1;
	pTab->get_ID(&nTabID);
	m_cmdManager.SetTabAndItemID(nTabID, nItemID);

	if( ::IsClipboardFormatAvailable(CF_TOOLBOX_ITEM) ||
		::IsClipboardFormatAvailable(CF_TEXT) ||
		::IsClipboardFormatAvailable(CF_UNICODETEXT))
	{
		menu.EnableMenuItem(ID_EDIT_PASTE, MF_BYCOMMAND|MF_ENABLED);
	}
}
	
 
LRESULT CToolbox::OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	if( m_wndRebar.IsWindow() )
	{
		m_wndRebar.MoveWindow(0, 0, LOWORD(lParam), HIWORD(lParam));
		
	}

	m_cmdManager.StopCmdProcessing();
	
	return 0;
}

STDMETHODIMP CToolbox::IOleObject_SetClientSite(LPOLECLIENTSITE pSite)
{
	HRESULT hr = CComControlBase::IOleObject_SetClientSite(pSite);
	// Check to see if the container has an ambient font. If it does,
	// clone it so your user can change the font of the control
	// without changing the ambient font for the container. If there is
	// no ambient font, create your own font object when you hook up a
	// client site.
	
	if(!m_pFont && pSite)
	{
		FONTDESC fd = g_FontDesc;
		CComPtr<IFont> pAF;
		CComPtr<IFont> pClone;
		if(SUCCEEDED(GetAmbientFont(&pAF)))
		{
			//clone the font
			if(SUCCEEDED(pAF->Clone(&pClone)))
				pClone->QueryInterface(IID_IFontDisp, (void**)&m_pFont);
			
		}
		else
		{
			hr = OleCreateFontIndirect(&fd,IID_IFontDisp,(void**)&m_pFont);
			if( FAILED(hr))
			{
				ATLASSERT(FALSE);
				return hr;
			}
		}
	}
	
	return hr;	
}


STDMETHODIMP CToolbox::get_ToolBoxTabs(IToolBoxTabs **pVal)
{
	CE_TBX_CHECK_PTR_ARG(pVal);
	return m_pTabs->QueryInterface(__uuidof(IToolBoxTabs), (void**)pVal);
}

STDMETHODIMP CToolbox::get_ActiveTab(IToolBoxTab **pVal)
{
	long nID=m_wndRebar.GetMaximizedBandID();
	CComVariant index(nID);
	return m_pTabs->get_Item(index, pVal);
}

STDMETHODIMP CToolbox::get_DTE(LPDISPATCH *pVal)
{
	CE_TBX_CHECK_PTR_ARG(pVal);
	*pVal = m_DTEDriver;

	return (*pVal)->AddRef();
}

STDMETHODIMP CToolbox::put_DTE(LPDISPATCH newVal)
{
	CE_TBX_CHECK_PTR_ARG(newVal);
	m_DTEDriver = newVal;
	return S_OK;
}

int CToolbox::InsertTab(int nIndex, BSTR bstrName)
{
	CString csName=bstrName;
	return m_wndRebar.InsertBand(nIndex, csName, -1);
}

BOOL CToolbox::DeleteTab(int TabID)
{
	return m_wndRebar.DeleteBand(m_wndRebar.IdToIndex(TabID));
}

void CToolbox::MaximizeTab(int TabID)
{
	m_wndRebar.MaximizeBand(m_wndRebar.IdToIndex(TabID));
}

CToolBarCtrl& CToolbox::GetToolBarCtrl(int TabID)
{
	return m_wndRebar.GetToolBarCtrl(m_wndRebar.IdToIndex(TabID));
}

BOOL CToolbox::ShowTab(int nTabID, BOOL bShow)
{
	BOOL bRet = m_wndRebar.ShowBand(m_wndRebar.IdToIndex(nTabID), bShow);
	return bRet;
}

STDMETHODIMP CToolbox::get_ShowPopupMenu(VARIANT_BOOL *pVal)
{
	CE_TBX_CHECK_PTR_ARG(pVal);
	*pVal=m_bShowPopupMenu;
	return S_OK;
}

STDMETHODIMP CToolbox::put_ShowPopupMenu(VARIANT_BOOL newVal)
{
	if( m_bShowPopupMenu != newVal )
	{
		m_bShowPopupMenu=newVal;
		SetDirty(true);
	}
	
	return S_OK;
}

void CToolbox::_setupFonts()
{
	if( m_pFont )
	{
		CComQIPtr<IFont> pFont = m_pFont;
		HFONT hFont=NULL;
		pFont->get_hFont(&hFont);
		if( m_wndRebar.IsWindow() )
			m_wndRebar.SetFont(m_pFont);

		if( m_wndEdit.IsWindow() )
			m_wndEdit.SetFont(hFont);
	}
}


void CToolbox::GetMassageBoxCaption(CString &csCaption)
{
	CComVariant varName;
	if( m_DTEDriver != NULL && 
		SUCCEEDED(m_DTEDriver.GetPropertyByName(L"Name", &varName)) )
	{
		if( varName.vt == VT_BSTR )
			csCaption = varName.bstrVal;
	}
	else
	{
		csCaption.LoadString(IDS_TOOLBOX);
	}
}

STDMETHODIMP CToolbox::get_PMICustomize(VARIANT_BOOL *pVal)
{
	CE_TBX_CHECK_PTR_ARG(pVal);
	*pVal=m_bPMICustomize;
	return S_OK;
}

STDMETHODIMP CToolbox::put_PMICustomize(VARIANT_BOOL newVal)
{
	if( m_bPMICustomize != newVal )
	{
		m_bPMICustomize=newVal;
		SetDirty(true);
	}
	
	return S_OK;
}

/**
*	Sets tab name in rebar ctrl and fires event  TabNameChanged
*	\param TabID - ID of tab
*	\param newVal - new Tab name
*/
STDMETHODIMP CToolbox::put_TabName(int TabID, BSTR newVal)
{
	CE_TBX_CHECK_PTR_ARG(newVal);
	CString csName=newVal;
	BOOL bRet = m_wndRebar.SetBandName(m_wndRebar.IdToIndex(TabID), csName);
	Fire_TabNameChanged(TabID, newVal);
	return bRet ? S_OK : E_FAIL;
}


/**
*	Returns next available comman id for toolbar button
*/
int CToolbox::GetNextCmdID()
{
	int nID = MIN_TBITEM_ID;
	for(int i=0; i<m_wndRebar.GetBandCount(); i++)
	{
		CToolBarCtrl& tb = m_wndRebar.GetToolBarCtrl(i);
		for(int j=0; j<tb.GetButtonCount(); j++)
		{
			TBBUTTON btn={0};
			tb.GetButton(j, &btn);
			if( nID < btn.idCommand )
			{
				nID = btn.idCommand;
			}
		}
	}

	return ++nID;
}

int CToolbox::GetItemWidth()
{
	return 22;
}
