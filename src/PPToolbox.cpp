// PPToolbox.cpp : Implementation of CPPToolbox
#include "stdafx.h"
#include "IfsToolbox.h"
#include "PPToolbox.h"

/////////////////////////////////////////////////////////////////////////////
// CPPToolbox

STDMETHODIMP CPPToolbox::Apply(void)
{
	if( IsPageDirty() == S_FALSE )
	{
		return S_OK;
	}
	
	SetDirty(FALSE);
	
	ATLTRACE(_T("CPPToolbox::Apply\n"));
	for (UINT i = 0; i < m_nObjects; i++)
	{
		CComQIPtr< IToolbox> pTlbx(m_ppUnk[i]);
		if( !pTlbx ) 
		{
			return E_NOINTERFACE;
		}
		
		CButton wndCheck = GetDlgItem(IDC_CHECK_SHOW_POPUP);
		pTlbx->put_ShowPopupMenu(wndCheck.GetCheck() == BST_CHECKED ? VARIANT_TRUE : VARIANT_FALSE);
		
		wndCheck = GetDlgItem(IDC_CHECK_CUSTOMIZE_ITEM);
		pTlbx->put_PMICustomize(wndCheck.GetCheck() == BST_CHECKED ? VARIANT_TRUE : VARIANT_FALSE);
	}
	
	return S_OK;
}

STDMETHODIMP CPPToolbox::Show(UINT nCmdShow) 
{	
	USES_CONVERSION;
	
	if(nCmdShow == SW_SHOW || nCmdShow == SW_SHOWNORMAL) 
	{
		for (UINT i = 0; i < m_nObjects; i++)
		{
			CComQIPtr< IToolbox> pTlbx(m_ppUnk[i]);
			if( !pTlbx )
			{
				return E_NOINTERFACE;
			}
			CComBool bShowPopup;
			pTlbx->get_ShowPopupMenu(&bShowPopup);
			CButton wndCheck = GetDlgItem(IDC_CHECK_SHOW_POPUP);
			wndCheck.SetCheck( bShowPopup ? BST_CHECKED : BST_UNCHECKED);
			
			CComBool bPMICustomize;
			pTlbx->get_PMICustomize(&bPMICustomize);
			wndCheck = GetDlgItem(IDC_CHECK_CUSTOMIZE_ITEM);
			wndCheck.SetCheck( bPMICustomize ? BST_CHECKED : BST_UNCHECKED);
		}
	}
	
	HRESULT hr = IPropertyPageImpl<CPPToolbox>::Show(nCmdShow);
		
	SetDirty(FALSE);
	return S_OK;
}
