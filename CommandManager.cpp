// CommandManager.cpp: implementation of the CCommandManager class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ifstoolbox.h"
#include "CommandManager.h"
#include "ifstoolbox.h"
#include "Toolbox.h"
#include "ToolBoxItem.h"
#include "ToolBoxItems.h"
#include "DataObjectEx.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

extern UINT CF_TOOLBOX_ITEM;
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCommandManager::CCommandManager(CToolbox& wndToolbox) : 
	m_wndTbx(wndToolbox),
	m_nActiveTabID(-1),
	m_nActiveItemID(-1),
	m_wndRebar(wndToolbox.m_wndRebar),
	m_wndEdit(wndToolbox.m_wndEdit)
{

}

void CCommandManager::SetTabAndItemID(int nTabID, int nItemID)
{
	ATLASSERT(nItemID>=MIN_TBITEM_ID||nItemID==-1); 
	m_nActiveItemID = nItemID; 
	m_nActiveTabID=nTabID;
};

CCommandManager::~CCommandManager()
{

}

inline void CCommandManager::OnAddNewTab()
{
	REBARBANDINFO rbi={0};
	rbi.cbSize = sizeof(rbi);
	rbi.fMask=RBBIM_ID;
	m_wndRebar.GetBandInfo(m_wndRebar.GetBandCount()-1, &rbi);
	
	CComPtr<IToolBoxTab> pTab;
	m_wndTbx.m_pTabs->GetTabByID(rbi.wID, &pTab);
	if( pTab)
	{
		pTab->Activate();
		pTab.Release();
	}
	
	CRect rcClient;
	m_wndTbx.GetClientRect(&rcClient);
	int nHeight=m_wndTbx.m_wndRebar.GetBarHeaderHeight();
	m_wndTbx.m_wndEdit.MoveWindow(1,rcClient.bottom-nHeight, rcClient.Width()-1, nHeight);
	m_wndTbx.m_wndEdit.SetWindowText(_T(""));
	m_wndTbx.m_wndEdit.ShowWindow(SW_SHOW);
	m_wndTbx.m_wndEdit.BringWindowToTop();
	m_wndTbx.m_wndEdit.SetFocus();
}

inline void CCommandManager::OnRenameTab()
{
	ATLASSERT(m_wndTbx.m_pTabs);
	CComPtr<IToolBoxTab> pTab;
	if( SUCCEEDED(m_wndTbx.m_pTabs->GetTabByID(m_nActiveTabID, &pTab)) )
	{
		CRect rc;
		CRect rcTmp;
		m_wndRebar.GetRect(m_wndRebar.IdToIndex(m_nActiveTabID), &rcTmp);

		rc.right = rcTmp.bottom;
		rc.left = rcTmp.top;
		rc.top = rcTmp.left;
		rc.bottom = rc.top + m_wndRebar.GetBarHeaderHeight();

		rc.InflateRect(-1, 0, 1,1);
				
		CComBSTR bstrName;
		pTab->get_Name(&bstrName);
		CString csName = bstrName;
		m_wndEdit.SetWindowText(csName);
		m_wndEdit.SendMessage(EM_SETSEL, 0, -1);

		m_wndEdit.MoveWindow(&rc);
		m_wndEdit.ShowWindow(SW_SHOW);
		m_wndEdit.BringWindowToTop();
		m_wndEdit.SetFocus();
	}	
}

inline void CCommandManager::OnDeleteTab()
{
	TCHAR lpszName[CAccessBarCtrl::_MaxButtonTextLength]={0};
	REBARBANDINFO rbi={0};
	rbi.cbSize = sizeof(rbi);
	rbi.fMask=RBBIM_ID|RBBIM_TEXT;
	rbi.cch = sizeof(lpszName);
	rbi.lpText=lpszName;
	if( m_wndRebar.GetBandInfo(m_wndRebar.IdToIndex(m_nActiveTabID), &rbi) )
	{
		CString csMsg;
		csMsg.Format(IDS_DELETE_TAB_Q_FMT, lpszName);
		CString csCaption;
		m_wndTbx.GetMassageBoxCaption(csCaption);
		if( m_wndTbx.MessageBox(csMsg, csCaption, MB_YESNO|MB_ICONQUESTION) == IDYES )
		{
			ATLVERIFY(m_wndTbx.m_pTabs->RemoveTab(m_nActiveTabID)==S_OK);				
		}
	}
}

inline void CCommandManager::OnMoveTabOrItem(bool bUp)
{
	if( m_nActiveItemID == -1 )
	{
		int nIndex = m_wndRebar.IdToIndex(m_nActiveTabID);
		int nNewIndex = !bUp ? nIndex+1 : nIndex-1;
		ATLVERIFY(m_wndRebar.MoveBand(nIndex, nNewIndex));
	}
	else
	{
		CToolBarCtrl& tb = m_wndTbx.GetToolBarCtrl(m_nActiveTabID);
		int nIndex = tb.CommandToIndex(m_nActiveItemID);
		ATLVERIFY(tb.MoveButton(nIndex, !bUp ? nIndex+1 : nIndex-1));
		for(int i=0; i<tb.GetButtonCount();i++)
		{
			TBBUTTON btn;
			tb.GetButton(i, &btn);
/*
			if( btn.idCommand == m_nActiveItemID )
				tb.HideButton(btn.idCommand, TRUE);
*/
		}
	}
	
}

inline void CCommandManager::OnHideTab()
{
	CComPtr<IToolBoxTab> pTab;
	if( SUCCEEDED(m_wndTbx.m_pTabs->GetTabByID(m_nActiveTabID, &pTab)) )
	{
		ATLVERIFY(pTab->put_Hidden(VARIANT_TRUE)==S_OK);
	}
}

inline void CCommandManager::OnToggleListView()
{
	CComPtr<IToolBoxTab> pTab;
	if( SUCCEEDED(m_wndTbx.m_pTabs->GetTabByID(m_nActiveTabID, &pTab)) )
	{
		CComBool listView;
		ATLVERIFY(pTab->get_ListView(&listView)==S_OK);
		listView = !listView;
		ATLVERIFY(pTab->put_ListView(listView)==S_OK);
	}
}

LRESULT CCommandManager::OnKillFocusEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	//StopCommandProcessing();
	return 0;
}

LRESULT CCommandManager::OnKeyDownEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
{
	int nVirtKey = (int) wParam; 
	long lKeyData = lParam;
	bHandled=FALSE;
	switch(nVirtKey) {
	case VK_ESCAPE:
		m_wndEdit.ShowWindow(SW_HIDE);
		m_wndEdit.SetWindowText(_T(""));
		bHandled=TRUE;
		break;
	case VK_RETURN:
		{
			CComBSTR bstNewName;
			m_wndEdit.GetWindowText(bstNewName.m_str);
			CString csCaption;
			m_wndTbx.GetMassageBoxCaption(csCaption);
			if( bstNewName.Length() == 0 )
			{
				CString csText;
				csText.LoadString(IDS_CAPTION_CANNOTBE_EMPTY);
				m_wndTbx.MessageBox(csText, csCaption, MB_OK|MB_ICONWARNING);
				m_wndEdit.SetFocus();
				break;
			}
			
			CComPtr<IToolBoxTab> pTab;
			CString csName=bstNewName;
			if( m_cmd == ID_ADD_TAB )
			{
				if( m_wndTbx.m_pTabs->TabWithNameExist(bstNewName) )
				{
					CString csText;
					csText.Format(IDS_E_TAB_WITHNAME_EXIST, csName);
					m_wndTbx.MessageBox(csText, csCaption, MB_OK|MB_ICONWARNING);
					m_wndEdit.SetFocus();
					break;
				}
				
				if( SUCCEEDED(m_wndTbx.m_pTabs->Add(bstNewName, &pTab)) )
				{
					ATLVERIFY(pTab->Activate()==S_OK);
				}
			}
			else if( m_cmd == ID_RENAME_TAB )
			{
				if( SUCCEEDED(m_wndTbx.m_pTabs->GetTabByID(m_nActiveTabID, &pTab)) )
				{
					CComBSTR bstrTmp;
					pTab->get_Name(&bstrTmp);
					if( !(bstrTmp == bstNewName) )
					{
						if( m_wndTbx.m_pTabs->TabWithNameExist(bstNewName) )
						{
							CString csText;
							csText.Format(IDS_E_TAB_WITHNAME_EXIST, csName);
							m_wndTbx.MessageBox(csText, csCaption, MB_OK|MB_ICONWARNING);
							m_wndEdit.SetFocus();
							break;
						}
						
						ATLVERIFY(SUCCEEDED(pTab->put_Name(bstNewName)));
					}
				}
			}
			else if( m_cmd == ID_RENAME_TABITEM )
			{
				CComPtr<IToolBoxItem> pItem;
				if( SUCCEEDED(m_wndTbx.m_pTabs->GetItemByID(m_nActiveItemID, &pItem)) )
				{
					CComBSTR bstrTmp;
					pItem->get_Name(&bstrTmp);
					if( !(bstrTmp == bstNewName) )
					{
						ATLVERIFY(SUCCEEDED(pItem->put_Name(bstNewName)));
					}
				}
			}
			
			m_wndEdit.ShowWindow(SW_HIDE);
			bHandled=TRUE;
		}
		break;
	default:
		break;
	}

	if( bHandled )
	{
		m_cmd=0;
	}
	
	return 0;
}

int CCommandManager::ProcessCmd(int cmd)
{
	bool bHandled=true;
	switch(cmd) {
	case ID_ADD_TAB:
		{
			OnAddNewTab();	
		}
		break;
	case ID_RENAME_TAB:
		{
			OnRenameTab();	
		}
		break;
	case ID_DELETE_TAB:
		{
			OnDeleteTab();
		}
		break;
	case ID_LIST_VIEW:
		{
			OnToggleListView();
		}
		break;
	case ID_HIDE_TAB:
		{
			OnHideTab();
		}
		break;
	case ID_SHOW_ALL_TABS:
		{
			m_wndTbx.m_pTabs->ShowAllTabs();
		}
		break;
	case ID_MOVE_DOWN:
		{
			OnMoveTabOrItem(false);
		}
		break;
	case ID_MOVE_UP:
		{
			OnMoveTabOrItem(true);
		}
		break;
	case ID_DELETE_TABITEM:
		{
			OnDeleteTabItem();
		}
		break;
	case ID_RENAME_TABITEM:
		{
			OnRenameTabItem();
		}
		break;
	case ID_EDIT_CUT:
		{
			OnEditCut();
		}
		break;
	case ID_EDIT_COPY:
		{
			OnEditCopy();
		}
		break;
	case ID_EDIT_PASTE:
		{
			OnEditPaste();
		}
		break;
	case ID_CUSTOMIZE_TBX:
		{
			CComPtr<IToolBoxTab> pTab;
			CComBSTR Name;
			m_wndTbx.m_pTabs->GetTabByID(m_nActiveTabID, &pTab);
			if( pTab )
			{
				pTab->get_Name(&Name);
			}
			
			m_wndTbx.Fire_CustomizeClicked(m_nActiveTabID, Name);
			
		}
		break;
	default:
		bHandled=false;
		StopCmdProcessing();
		break;
	}

	if(bHandled)
	{
		m_cmd = cmd;
	}
	
	return 0;
}

inline void CCommandManager::OnDeleteTabItem()
{
	ATLASSERT(m_nActiveItemID>=MIN_TBITEM_ID);
	CComPtr<IToolBoxItem> pItem;
	m_wndTbx.m_pTabs->GetItemByID(m_nActiveTabID, m_nActiveItemID, &pItem);
	ATLASSERT(pItem);
	if( pItem )
	{
		CComBSTR Name;
		pItem->get_Name(&Name);
		CString csName = Name;
		CString csMsg;
		csMsg.Format(IDS_DELETE_ITEM_Q_FMT, csName);
		CString csCaption;
		m_wndTbx.GetMassageBoxCaption(csCaption);
		if( m_wndTbx.MessageBox(csMsg, csCaption, MB_YESNO|MB_ICONQUESTION) == IDYES )
		{
			ATLVERIFY(SUCCEEDED(pItem->Delete()));		
		}
	}

}

void CCommandManager::StopCmdProcessing()
{
	if( m_wndEdit.IsWindowVisible() )
	{
		m_wndEdit.ShowWindow(SW_HIDE);
		m_wndEdit.SetWindowText(_T(""));
	}
}

void CCommandManager::OnRenameTabItem()
{
	ATLASSERT(m_wndTbx.m_pTabs);

	CComObject<CToolBoxTab>* pTab = m_wndTbx.m_pTabs->GetTabByID(m_nActiveTabID);
	ADDREF(pTab);
	if( pTab )
	{
		ATLVERIFY(SUCCEEDED(pTab->put_ListView(VARIANT_TRUE)));

		CComPtr<IToolBoxItem> pItem;
		ATLVERIFY(SUCCEEDED(pTab->GetItem(m_nActiveItemID, &pItem)));
		
		CToolBarCtrl& tb = m_wndRebar.GetBandData(m_wndRebar.IdToIndex(m_nActiveTabID)).ctrlToolBar;
		CRect rc;
		ATLVERIFY(tb.GetRect(m_nActiveItemID, &rc));
		ATLVERIFY(MapWindowPoints(tb, m_wndRebar, (LPPOINT)&rc, 2));
		rc.InflateRect(-1, 0, 1,1);		
		CComBSTR bstrName;
		pItem->get_Name(&bstrName);
		CString csName = bstrName;
		ATLVERIFY(m_wndEdit.SetWindowText(csName));
		m_wndEdit.SendMessage(EM_SETSEL, 0, -1);
		
		m_wndEdit.MoveWindow(&rc);
		m_wndEdit.ShowWindow(SW_SHOW);
		m_wndEdit.BringWindowToTop();
		m_wndEdit.SetFocus();
	}
	RELEASE(pTab);
}

void CCommandManager::OnEditCopy()
{
	ATLASSERT(m_nActiveItemID>=MIN_TBITEM_ID);
	CComPtr<IToolBoxItem> pItem;
	m_wndTbx.m_pTabs->GetItemByID(m_nActiveTabID, m_nActiveItemID, &pItem);
	ATLASSERT(pItem);
	if( pItem )
	{
		CComQIPtr<IPersistStreamInit> pPSI = pItem;
		CComObject<CDataObjectImpl>* pDO=NULL;
		CComObject<CDataObjectImpl>::CreateInstance(&pDO);

		IStream* pStream=NULL;
		ATLVERIFY(SUCCEEDED(::CreateStreamOnHGlobal(NULL, TRUE, &pStream)));
		ATLVERIFY(SUCCEEDED(pPSI->Save(pStream, FALSE)));
		ATLVERIFY(SUCCEEDED(pDO->SetIStreamData(CF_TOOLBOX_ITEM, pStream, TRUE)));		
		ATLVERIFY(SUCCEEDED(::OleSetClipboard(pDO)));
	}	
}

void CCommandManager::OnEditCut()
{
	ATLASSERT(m_nActiveItemID>=MIN_TBITEM_ID);
	CComPtr<IToolBoxItem> pItem;
	m_wndTbx.m_pTabs->GetItemByID(m_nActiveTabID, m_nActiveItemID, &pItem);
	ATLASSERT(pItem);
	if( pItem )
	{
		CComBSTR Name;
		pItem->get_Name(&Name);
		CString csName = Name;
		CString csMsg;
		csMsg.Format(IDS_DELETE_ITEM_Q_FMT, csName);
		CString csCaption;
		m_wndTbx.GetMassageBoxCaption(csCaption);
		if( m_wndTbx.MessageBox(csMsg, csCaption, MB_YESNO|MB_ICONQUESTION) == IDYES )
		{
			OnEditCopy();
			ATLVERIFY(SUCCEEDED(pItem->Delete()));		
		}
	}	
}

void CCommandManager::OnEditPaste()
{
	CComObject<CToolBoxTab>* pTab;
	pTab = m_wndTbx.m_pTabs->GetTabByID(m_nActiveTabID);
	ATLASSERT(pTab);
	ADDREF(pTab);
	if( ::IsClipboardFormatAvailable(CF_TOOLBOX_ITEM) )
	{
		// copy TBx item
		CComPtr<IDataObject> pDO;
		ATLVERIFY(SUCCEEDED(::OleGetClipboard(&pDO)));
		FORMATETC fe;
		SETFORMATETC(fe, CF_TOOLBOX_ITEM, DVASPECT_CONTENT, NULL, TYMED_ISTREAM, -1);
		STGMEDIUM med={0};
		ATLVERIFY(SUCCEEDED(pDO->GetData(&fe, &med)));
		if( med.pstm )
		{
			CComObject<CToolBoxItem>* pItem;
			CComObject<CToolBoxItem>::CreateInstance(&pItem);
			ADDREF(pItem);
			if( pItem )
			{
				ATLVERIFY(SUCCEEDED(pItem->put_Collection(pTab->GetItems())));
				CComQIPtr<IPersistStreamInit> pPSI = pItem;
				LARGE_INTEGER dlibMove = { 0, 0 };
				ATLVERIFY(SUCCEEDED(med.pstm->Seek(dlibMove, STREAM_SEEK_SET, NULL)));
				ATLVERIFY(SUCCEEDED(pPSI->Load(med.pstm)));
				CComVariant var;
				pTab->GetItems()->Add(pItem, NULL, var);
				pItem->Select();
			}
		}

		::ReleaseStgMedium(&med);
	}
	else if( ::IsClipboardFormatAvailable(CF_TEXT) || ::IsClipboardFormatAvailable(CF_UNICODETEXT) )
	{
		BOOL bText = ::IsClipboardFormatAvailable(CF_TEXT);
		CComPtr<IDataObject> pDO;
		ATLVERIFY(SUCCEEDED(::OleGetClipboard(&pDO)));
		FORMATETC fe;
		SETFORMATETC(fe, bText ? CF_TEXT : CF_UNICODETEXT, DVASPECT_CONTENT, NULL, TYMED_HGLOBAL, -1);
		STGMEDIUM med={0};
		ATLVERIFY(SUCCEEDED(pDO->GetData(&fe, &med)));
		if( med.hGlobal )
		{
			CComBSTR Text;
			if( bText )
			{
				// we asked for a global memory handle, must lock it for access
				LPCTSTR ptxt = (LPCTSTR)GlobalLock(med.hGlobal);
				Text = ptxt;
				GlobalUnlock(med.hGlobal);
			}
			else
			{
				// we asked for a global memory handle, must lock it for access
				LPWSTR ptxt = (LPWSTR)GlobalLock(med.hGlobal);
				Text = ptxt;
				GlobalUnlock(med.hGlobal);
			}
			
			CComObject<CToolBoxItem>* pItem;
			CComObject<CToolBoxItem>::CreateInstance(&pItem);
			ADDREF(pItem);
			if( pItem )
			{
				ATLVERIFY(SUCCEEDED(pItem->put_Collection(pTab->GetItems())));
				pItem->put_Format(ToolBoxItemFormatText);
				CComVariant Data = Text;
				pTab->GetItems()->Add(pItem, NULL, Data);
				pItem->Select();
			}
			
		}

		::ReleaseStgMedium(&med);
	}

	RELEASE(pTab);
}

int CCommandManager::GetTabID()
{
	return m_nActiveTabID;
}

int CCommandManager::GetItemID()
{
	return m_nActiveItemID;
}
