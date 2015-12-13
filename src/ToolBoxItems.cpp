// ToolBoxItems.cpp : Implementation of CToolBoxItems
#include "stdafx.h"
#include "IfsToolbox.h"
#include "ToolBoxItems.h"
#include "ToolBoxTab.h"
#include "ToolBoxTabs.h"
#include "Toolbox.h"
#include "Guid.h"

#define CE_TBIS_CHECK_PTR(ptr, uiResStr, error) if( !(ptr) ) {  return Error(uiResStr, IID_IToolBoxItems, error); }
#define CE_TBIS_CHECK_BOOL(b, uiResStr, error) if( !(b) ) {  return Error(uiResStr, IID_IToolBoxItems, error); }
#define CE_TBIS_CHECK_HR(hr, uiResStr) CE_CHECK_HR(hr, IID_IToolBoxItems, uiResStr)
#define CE_TBIS_CHECK_PTR_ARG(ptr) if( !(ptr) ) {  return Error(IDS_E_INVALID_ARG, IID_IToolBoxItems, E_INVALIDARG); }


HRESULT GetCtrlBitmap(const CLSID& clsid, const IID& iid, HINSTANCE hInst, BSTR bstrGuid, CBitmapHandle& bmp)
{
	TCHAR   szValue[_MAX_PATH] ={0};
	ULONG dwCount=sizeof(szValue)-2;
	CGuid guid=bstrGuid;
	CString	csCLSIDKey;
	csCLSIDKey.Format(_T("CLSID\\%s\\ToolboxBitmap32"), (LPCTSTR)(CString)guid);
	CRegKey rkCtrl;
	LONG lRet = rkCtrl.Open(HKEY_CLASSES_ROOT, csCLSIDKey);
	if( lRet != ERROR_SUCCESS )
	{
		csCLSIDKey.Format(_T("CLSID\\{%s}\\ToolboxBitmap"), guid);
		lRet = rkCtrl.Open(HKEY_CLASSES_ROOT, csCLSIDKey);
		return AtlReportError(clsid, IDS_E_CANNOT_OPEN_RK_FOR_CTRL_FMT, 0, NULL, iid, REGDB_E_CLASSNOTREG, hInst);
	}
	
	if( rkCtrl.QueryValue(szValue, _T(""), &dwCount) != ERROR_SUCCESS )
	{
		return AtlReportError(clsid, IDS_E_CANNOT_OPEN_RK_FOR_CTRL_FMT, 0, NULL, iid, REGDB_E_KEYMISSING, hInst);
	}
	
	TCHAR * pszBitmapModule = _tcstok(szValue, _T(","));
	TCHAR * pszBitmapId = _tcstok(NULL, _T(","));
	HINSTANCE hinstBitmap;
	int nBitmapId;
	if ((pszBitmapModule != NULL) && (pszBitmapId != NULL))
	{
		LPTSTR p ;
		if (*pszBitmapModule == '"')
		{
			lstrcpy( pszBitmapModule, pszBitmapModule+1 ) ;
			p = strrchr( pszBitmapModule, '"' ) ;
			if (p) *p = '\0' ;
		}
		nBitmapId = _ttoi(pszBitmapId);
		
		hinstBitmap = ::LoadLibrary(pszBitmapModule);
		
		if (hinstBitmap && (nBitmapId > 0))
		{
			bmp = ::LoadBitmap(hinstBitmap, MAKEINTRESOURCE(nBitmapId));
			if(!bmp) //If the ToolBoxBitmap32 does not have a Bitmap we'll give one. We dont accept icons as of this version.
			{
				bmp = ::LoadBitmap(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDB_CTRL_DEFAULT));
			}
			//cleanup
			ATLASSERT(::FreeLibrary(hinstBitmap));
		}
	}
	
	return S_OK;
}
/////////////////////////////////////////////////////////////////////////////
// CToolBoxItems

STDMETHODIMP CToolBoxItems::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IToolBoxItems
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CToolBoxItems::FinalConstruct()
{		
	return S_OK;
}

STDMETHODIMP CToolBoxItems::FinalRelease()
{
	HRESULT hr=S_OK;
	
	std::for_each(m_vecItems.begin(), m_vecItems.end(), releaseItem);
	
	m_vecItems.clear();
	
	
	return hr;
}

STDMETHODIMP CToolBoxItems::get__NewEnum(LPUNKNOWN *pVal)
{
	CE_TBIS_CHECK_PTR_ARG(pVal);
	
	typedef CComEnumOnSTL<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT,
		_CopyIDispatch, VectorItems > CComEnumUnknownOnVector;
	
	return CreateSTLEnumerator<CComEnumUnknownOnVector>(pVal, (IToolBoxItems*)this, m_vecItems);
}

STDMETHODIMP CToolBoxItems::get_Item(VARIANT index, IToolBoxItem **pVal)
{
	CE_TBIS_CHECK_PTR_ARG(pVal);

	if( !(index.vt==VT_BSTR||index.vt==VT_I4) ) 
	{  
		return Error(IDS_E_INVALID_ARG, __uuidof(IToolBoxItems), E_INVALIDARG);
	}
	
	if( index.vt==VT_I4 )
	{
		if( !(index.lVal>=0 && index.lVal<m_vecItems.size()) ) 
		{  
			return Error(IDS_E_INVALID_INDEX, __uuidof(IToolBoxItems), DISP_E_BADINDEX);
		}
		
		return m_vecItems[index.lVal].pItem->QueryInterface(__uuidof(IToolBoxItem), (void**)pVal);
	}
	
	if( index.vt==VT_BSTR )
	{
		CE_TBIS_CHECK_PTR_ARG(index.bstrVal);
		vecItemIter it = std::find_if(m_vecItems.begin(), m_vecItems.end(), findItemByName(index.bstrVal));
		if( it != m_vecItems.end() )
		{
			return (*it).pItem->QueryInterface(__uuidof(IToolBoxItem), (void**)pVal);
		}
	}
	
	return Error(IDS_E_INVALID_INDEX, __uuidof(IToolBoxItems), DISP_E_BADINDEX); 
}

STDMETHODIMP CToolBoxItems::get_Count(long *pVal)
{
	CE_TBIS_CHECK_PTR_ARG(pVal);
	*pVal=m_vecItems.size();
	return S_OK;
}

STDMETHODIMP CToolBoxItems::get_SelectedItem(IToolBoxItem **pVal)
{
	CE_TBIS_CHECK_PTR_ARG(pVal);

	ATLASSERT(m_pTab);
	*pVal=NULL;
	for(int i=0; i<m_vecItems.size(); i++)
	{
		if( m_vecItems[i].pItem->IsSelected() )
		{
			return m_vecItems[i].pItem->QueryInterface(__uuidof(IToolBoxItem), (void**)pVal);
		}
	}
		
	return E_FAIL;
}

STDMETHODIMP CToolBoxItems::get_Parent(IToolBoxTab **pVal)
{
	CE_TBIS_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pTab);
	return m_pTab->QueryInterface(__uuidof(IToolBoxTab), (void**)pVal);
}

STDMETHODIMP CToolBoxItems::Add(BSTR Name, VARIANT Data, ToolBoxItemFormat Format, IToolBoxItem **ppItem)
{
	CE_TBIS_CHECK_PTR_ARG(Name);
	CComObject<CToolBoxItem>* pItem=NULL;
	HRESULT hr = CComObject<CToolBoxItem>::CreateInstance(&pItem);
	ADDREF(pItem);
	pItem->put_Collection(this);
	pItem->put_Format(Format);
	pItem->put_Data(Data);
	pItem->put_Name(Name);
	hr = Add(pItem, Name, Data);
	if( ppItem )
	{
		return pItem->QueryInterface(ppItem);
	}		

	return hr;
}

void CToolBoxItems::AddToolBarButton(int nCmdID, int nIndex, int nBitmapIdx, const CString& csName)
{
	ATLASSERT(m_pTab!=NULL);
	CToolBarCtrl& wndTlbr = m_pTab->GetToolBarCtrl();

	TBBUTTON tb={0};
	tb.iBitmap = nBitmapIdx; 
	tb.idCommand = nCmdID; 
	tb.fsState = TBSTATE_ENABLED;
	tb.fsStyle = BYTE(TBSTYLE_GROUP|TBSTYLE_WRAPABLE|TBSTYLE_BUTTON);
	tb.dwData = 0; 
	tb.iString = (int)(LPTSTR)(LPCTSTR)csName;
	
	ATLVERIFY(wndTlbr.InsertButton(nIndex, &tb));

	CComBool bListView;
	m_pTab->get_ListView(&bListView);
	
	CRect rc;
	m_pTab->GetToolbox()->GetClientRect(&rc);
	TBBUTTONINFO tbi={0};
	tbi.cbSize=sizeof(TBBUTTONINFO);
	tbi.dwMask = TBIF_SIZE;
	tbi.cx = bListView ? rc.Width()-2 : m_pTab->GetToolbox()->GetItemWidth();

	ATLVERIFY(wndTlbr.SetButtonInfo(tb.idCommand, &tbi));

	RECT rc1;
	if( bListView )
	{
		wndTlbr.SetRows(wndTlbr.GetButtonCount(), FALSE, &rc1);
	}
	else
	{
		wndTlbr.SetRows(1, TRUE, &rc1);
	}
}

STDMETHODIMP CToolBoxItems::get_DTE(LPDISPATCH *pVal)
{
	CE_TBIS_CHECK_PTR_ARG(pVal);
	ATLASSERT(m_pTab);
	return m_pTab->get_DTE(pVal);
}

void CToolBoxItems::put_Parent(CToolBoxTab *pTab)
{
	ATLASSERT(pTab);
	m_pTab=pTab;
}

CToolBoxTab* CToolBoxItems::GetTab()
{
	return m_pTab;
}

HRESULT CToolBoxItems::SelectItem(UINT nItemID)
{
	ATLASSERT(m_pTab);
	CE_TBIS_CHECK_BOOL(nItemID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	vecItemIter it = std::find_if(m_vecItems.begin(), m_vecItems.end(), findItemByID(nItemID));
	if( it != m_vecItems.end() )
	{
		CComObject<CToolBoxItem>* pItem=(*it).pItem;
		CComBSTR bstrName;
		pItem->get_Name(&bstrName);

		CToolBarCtrl& wndTlbr=m_pTab->GetToolBarCtrl();
		if( !wndTlbr.IsButtonChecked(nItemID) )
		{
			wndTlbr.CheckButton(nItemID, TRUE);			
			m_pTab->GetToolbox()->Fire_ItemSelected(nItemID, pItem);
		}
		return S_OK;
	}
	
	
	return Error(IDS_E_INVALID_INDEX, IID_IToolBoxItems, DISP_E_BADINDEX);
}

HRESULT CToolBoxItems::HideItem(UINT nItemID,bool bHide)
{
	ATLASSERT(m_pTab);
	CE_TBIS_CHECK_BOOL(nItemID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	vecItemIter it = std::find_if(m_vecItems.begin(), m_vecItems.end(), findItemByID(nItemID));
	if( it != m_vecItems.end() )
	{
		CComObject<CToolBoxItem>* pItem=(*it).pItem;
		CComBSTR bstrName;
		pItem->get_Name(&bstrName);
		
		CToolBarCtrl& wndTlbr=m_pTab->GetToolBarCtrl();
		wndTlbr.HideButton(nItemID, bHide);
		
		m_pTab->GetToolbox()->Fire_ItemHidden(nItemID, bstrName);
		return S_OK;
	}
	
	
	return Error(IDS_E_INVALID_INDEX, IID_IToolBoxItems, DISP_E_BADINDEX);
}

HRESULT CToolBoxItems::GetItem(int nItemID, IToolBoxItem** ppItem)
{
	CE_TBIS_CHECK_PTR_ARG(ppItem);
	CE_TBIS_CHECK_BOOL(nItemID>=MIN_TBITEM_ID, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	ATLASSERT(m_pTab);
	vecItemIter it = std::find_if(m_vecItems.begin(), m_vecItems.end(), findItemByID(nItemID));
	if( it != m_vecItems.end() )
	{
		return (*it).pItem->QueryInterface(__uuidof(IToolBoxItem), (void**)ppItem);
	}
	
	
	return Error(IDS_E_INVALID_INDEX, IID_IToolBoxItems, DISP_E_BADINDEX);
}

HRESULT CToolBoxItems::GetSelectedItem(IToolBoxItem** ppItem)
{
	CE_TBIS_CHECK_PTR_ARG(ppItem);
	ATLASSERT(m_pTab);
	vecItemIter it = std::find_if(m_vecItems.begin(), m_vecItems.end(), findSelectedItem());
	if( it != m_vecItems.end() )
	{
		return (*it).pItem->QueryInterface(__uuidof(IToolBoxItem), (void**)ppItem);
	}
	
	
	return Error(IDS_E_INVALID_INDEX, IID_IToolBoxItems, DISP_E_BADINDEX);
}

HRESULT CToolBoxItems::DeleteItem(UINT nItemID)
{
	ATLASSERT(m_pTab);
	CE_TBIS_CHECK_BOOL(nItemID>=0, IDS_E_INVALID_INDEX, DISP_E_BADINDEX);
	vecItemIter it = std::find_if(m_vecItems.begin(), m_vecItems.end(), findItemByID(nItemID));
	if( it != m_vecItems.end() )
	{
		CComObject<CToolBoxItem>* pItem=(*it).pItem;
		
		CToolBarCtrl& wndTlbr=m_pTab->GetToolBarCtrl();
		wndTlbr.DeleteButton(wndTlbr.CommandToIndex(nItemID));

		CComBSTR bstrName;
		pItem->get_Name(&bstrName);

		m_vecItems.erase(it);
		RELEASE(pItem);
		
		m_pTab->GetToolbox()->Fire_ItemRemoved(nItemID, bstrName);
		return S_OK;
	}
	
	
	return Error(IDS_E_INVALID_INDEX, IID_IToolBoxItems, DISP_E_BADINDEX);
}

bool CToolBoxItems::ItemWithNameExist(BSTR bstrName)
{
	return std::find_if(m_vecItems.begin(), m_vecItems.end(), findItemByName(bstrName)) != m_vecItems.end();
}

bool CToolBoxItems::IsItemSelected(int nItemID)
{
	ATLASSERT(m_pTab);
	return m_pTab->GetToolBarCtrl().IsButtonChecked(nItemID) ? true : false;
}

HRESULT CToolBoxItems::Add(CComObject<CToolBoxItem> *pItem, BSTR Name, VARIANT Data)
{
	if( !pItem || !m_pTab )
	{
		ATLASSERT(FALSE);
		return E_POINTER;
	}

	HRESULT hr=S_OK;

	int nItemIndex=-1;
	int nImageIndex = -1;
	int nItemID=-1;
	CComVariant DataTmp;
	CComBSTR bstrDescription;
	CComBSTR bstrName;
	ToolBoxItemFormat Format;

	CToolBarCtrl& wndTlbr=m_pTab->GetToolBarCtrl();
	
	pItem->get_Format(&Format);	
	pItem->get_Description(&bstrDescription);
	

	CString csName;

	switch(Format) 
	{
	case ToolBoxItemFormatText:
	case ToolBoxItemFormatHTML:
		{
			if( Data.vt == VT_EMPTY )
			{
				pItem->get_Data(&Data);
			}

			CE_TBIS_CHECK_PTR(Data.vt == VT_BSTR && Data.bstrVal, IDS_E_INVALID_ARG, E_INVALIDARG);
					
			if( !bstrDescription.Length() )
			{
				bstrDescription.LoadString(Format==ToolBoxItemFormatText ? IDS_TEXT:IDS_HTML);				
				pItem->put_Description(bstrDescription);
			}

			csName=bstrDescription;
			csName += _T(": ");
			csName += Data.bstrVal;
			if( csName.GetLength() > 100 )
			{
				csName = csName.Left(50);
			}

			bstrName = csName;
			pItem->put_Name(bstrName);
			
			nImageIndex = Format==ToolBoxItemFormatText ? IMG_IDX_TEXT:IMG_IDX_HTML;
			nItemID = m_pTab->GetToolbox()->GetNextCmdID();
		}
		break;
	case ToolBoxItemFormatPointer:
		{
			CComPtr<IToolBoxItem> pItemTemp;
			GetItem(POINTER_ID, &pItemTemp);
			if( pItemTemp )
			{
				CString csErr;
				csErr.LoadString(IDS_E_ITEM_POINTER_ALREADY_EXIST);
				return Error(csErr, __uuidof(IToolBoxItems), E_INVALIDARG);
			}

			if( !bstrDescription.Length() )
			{
				bstrDescription.LoadString(IDS_POINTER);
				pItem->put_Description(bstrDescription);
			}

			if( Name )
			{
				csName = Name;
			}
			else
			{
				pItem->get_Name(&bstrName);
				csName = bstrName;
			}

			
			nItemIndex = 0;
			nItemID = POINTER_ID;
			nImageIndex = IMG_IDX_POINTER;				
		}
		break;
	case ToolBoxItemFormatGUID:
		{
			if( Data.vt == VT_EMPTY )
			{
				pItem->get_Data(&Data);
			}
			
			CE_TBIS_CHECK_PTR(Data.vt == VT_BSTR && Data.bstrVal, IDS_E_INVALID_ARG, E_INVALIDARG);
			CBitmapHandle bmp;
			hr = GetCtrlBitmap(GetObjectCLSID(), __uuidof(IToolBoxItems), _Module.GetModuleInstance(), Data.bstrVal, bmp);
			if( SUCCEEDED(hr) )
			{
				ATLVERIFY(wndTlbr.AddBitmap(1, bmp)!=-1);
				nImageIndex = wndTlbr.GetButtonCount() < m_nBitmapAdded ? m_nBitmapAdded : ++m_nBitmapAdded;

				if( !bstrDescription.Length() )
				{					
					CString csProgID;
					CGuid guid=Data.bstrVal;
					guid.ProgID(csProgID);
					int nFirstIdx=csProgID.Find(TCHAR('.'))+1;
					int nSecondIdx=csProgID.Find(TCHAR('.'), nFirstIdx);
					bstrDescription = csProgID.Mid(nFirstIdx, nSecondIdx-nFirstIdx);
					pItem->put_Description(bstrDescription);
				}

				if( Name )
				{
					csName = Name;
				}
				else
				{
					pItem->get_Name(&bstrName);
					csName = bstrName;
				}
				
				nItemID = m_pTab->GetToolbox()->GetNextCmdID();	
			}
		}
		break;
	default:
		return Error(IDS_E_INVALID_ARG, IID_IToolBoxItems, E_INVALIDARG);
		break;
	}

	pItem->put_ID(nItemID);
	pItem->put_Data(Data);

	AddToolBarButton(nItemID, nItemIndex, nImageIndex, csName);
	
	m_vecItems.push_back(_tHolder(pItem));

	m_pTab->GetToolbox()->Fire_ItemAdded(nItemID, bstrName);

	return S_OK;
}
