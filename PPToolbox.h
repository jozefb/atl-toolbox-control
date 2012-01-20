// PPToolbox.h : Declaration of the CPPToolbox

#ifndef __PPTOOLBOX_H_
#define __PPTOOLBOX_H_

#include "resource.h"       // main symbols

EXTERN_C const CLSID CLSID_PPToolbox;

/////////////////////////////////////////////////////////////////////////////
// CPPToolbox
class ATL_NO_VTABLE CPPToolbox :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CPPToolbox, &CLSID_PPToolbox>,
	public IPropertyPageImpl<CPPToolbox>,
	public CDialogImpl<CPPToolbox>
{
public:
	CPPToolbox() 
	{
		m_dwTitleID = IDS_TITLEPPToolbox;
		m_dwHelpFileID = IDS_HELPFILEPPToolbox;
		m_dwDocStringID = IDS_DOCSTRINGPPToolbox;
	}

	enum {IDD = IDD_PPTOOLBOX};

DECLARE_REGISTRY_RESOURCEID(IDR_PPTOOLBOX)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CPPToolbox) 
	COM_INTERFACE_ENTRY(IPropertyPage)
END_COM_MAP()

BEGIN_MSG_MAP(CPPToolbox)
	CHAIN_MSG_MAP(IPropertyPageImpl<CPPToolbox>)
	COMMAND_HANDLER(IDC_CHECK_SHOW_POPUP, BN_CLICKED, OnClickedShowPopup)
	COMMAND_HANDLER(IDC_CHECK_CUSTOMIZE_ITEM, BN_CLICKED, OnClickedCustomizeItem)
END_MSG_MAP() 
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	STDMETHOD(Apply)(void);
	STDMETHOD(Show)(UINT nCmdShow);
	LRESULT OnClickedShowPopup(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		SetDirty(true);
		CButton btnShowPopup = hWndCtl;
		
		bool bShowPopup = btnShowPopup.GetCheck() == BST_CHECKED;
		::EnableWindow(GetDlgItem(IDC_STATIC_PMI), bShowPopup);
		::EnableWindow(GetDlgItem(IDC_CHECK_CUSTOMIZE_ITEM), bShowPopup);
		
		return 0;
	}
	LRESULT OnClickedCustomizeItem(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		SetDirty(true);
		return 0;
	}
};

#endif //__PPTOOLBOX_H_
