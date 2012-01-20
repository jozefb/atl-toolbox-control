// Toolbox.h : Declaration of the CToolbox

#ifndef __TOOLBOX_H_
#define __TOOLBOX_H_

#include "resource.h"       // main symbols
#include <atlctl.h>
#include "AccessBar.h"
#include "ToolBoxTabs.h"
#include "IfsToolboxCP.h"
#include "atlctrlw.h"
#include "AtlCtrlXp.h"

#include "CommandManager.h"	// Added by ClassView

typedef enum {
	POINTER_ID=100,
	MIN_TBITEM_ID = POINTER_ID,
	
};
/////////////////////////////////////////////////////////////////////////////
// CToolbox
class ATL_NO_VTABLE CToolbox : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CStockPropImpl<CToolbox, IToolbox, &IID_IToolbox, &LIBID_IFSTOOLBOXLib>,
	public CComControl<CToolbox>,
	public IPersistStorageImpl<CToolbox>,
	public IPersistPropertyBagImpl<CToolbox>,
	public IPersistStreamInitImpl<CToolbox>,
	public IOleControlImpl<CToolbox>,
	public IOleObjectImpl<CToolbox>,
	public IOleInPlaceActiveObjectImpl<CToolbox>,
	public IViewObjectExImpl<CToolbox>,
	public IOleInPlaceObjectWindowlessImpl<CToolbox>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CToolbox>,
	public ISpecifyPropertyPagesImpl<CToolbox>,
	public IQuickActivateImpl<CToolbox>,
	public IDataObjectImpl<CToolbox>,
	public IProvideClassInfo2Impl<&CLSID_Toolbox, &DIID_IToolboxEvents, &LIBID_IFSTOOLBOXLib>,
	public IPropertyNotifySinkCP<CToolbox>,
	public CComCoClass<CToolbox, &CLSID_Toolbox>,
	public CProxyIToolboxEvents< CToolbox >,
	public IObjectSafetyImpl<CToolbox, INTERFACESAFE_FOR_UNTRUSTED_CALLER |INTERFACESAFE_FOR_UNTRUSTED_DATA>
	
{
public:
	CToolbox();

	friend class CCommandManager;

DECLARE_REGISTRY_RESOURCEID(IDR_TOOLBOX)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CToolbox)
	COM_INTERFACE_ENTRY(IToolbox)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IViewObjectEx)
	COM_INTERFACE_ENTRY(IViewObject2)
	COM_INTERFACE_ENTRY(IViewObject)
	COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceObject)
	COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
	COM_INTERFACE_ENTRY(IOleControl)
	COM_INTERFACE_ENTRY(IOleObject)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(IPersistPropertyBag)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
	COM_INTERFACE_ENTRY(IQuickActivate)
	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY(IDataObject)
	COM_INTERFACE_ENTRY(IProvideClassInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)
	COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
	COM_INTERFACE_ENTRY(IObjectSafety)	
END_COM_MAP()

BEGIN_PROP_MAP(CToolbox)
	PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
	PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
	PROP_ENTRY("Enabled", DISPID_ENABLED, CLSID_NULL)
	PROP_ENTRY("Font", DISPID_FONT, CLSID_StockFontPage)
	PROP_ENTRY("ShowPopupMenu", 4, CLSID_NULL)
	PROP_ENTRY("PMCustomizeItem", 5, CLSID_NULL)
	PROP_PAGE(CLSID_PPToolbox)
	// Example entries
	// PROP_ENTRY("Property Description", dispid, clsid)
	// PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()

BEGIN_CONNECTION_POINT_MAP(CToolbox)
	CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
	CONNECTION_POINT_ENTRY(DIID_IToolboxEvents)
END_CONNECTION_POINT_MAP()

BEGIN_MSG_MAP(CToolbox)
	MESSAGE_HANDLER(WM_CREATE, OnCreate)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
	MESSAGE_HANDLER(WM_COMMAND, OnCommand)
	MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
	MESSAGE_HANDLER(WM_CTLCOLOREDIT,OnCtlColorEdit)	
	MESSAGE_HANDLER(WM_LBUTTONDOWN,OnLButtonDown)
	MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCode)
	CHAIN_MSG_MAP(CComControl<CToolbox>)
	DEFAULT_REFLECTION_HANDLER()
	NOTIFY_CODE_HANDLER(TTN_NEEDTEXT, OnRelayToolTip)
	NOTIFY_CODE_HANDLER(RBN_LAYOUTCHANGED, OnRebarLayoutChanged);
	REFLECT_NOTIFICATIONS()	
ALT_MSG_MAP(1)
	MESSAGE_HANDLER(WM_CHAR, OnCharEdit)
	MESSAGE_HANDLER(WM_KEYDOWN, m_cmdManager.OnKeyDownEdit)
	MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCodeEdit)
	MESSAGE_HANDLER(EN_KILLFOCUS, m_cmdManager.OnKillFocusEdit)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);



// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IViewObjectEx
	DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

// IQuickActivate

// IPersistStreamInit
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty);
	STDMETHOD(Load)(LPSTREAM pStm);

// IOleControl
	STDMETHOD(OnAmbientPropertyChange)(DISPID dispid);

	STDMETHOD(IOleObject_SetClientSite)(LPOLECLIENTSITE pSite);

	STDMETHOD(FinalConstruct)();
	STDMETHOD(FinalRelease)();

	HRESULT FireOnChanged(DISPID dispID);
// IToolbox
public:
	int GetItemWidth();
	int GetNextCmdID();
	STDMETHOD(put_TabName)(/*[in]*/ int TabID, /*[in]*/ BSTR newVal);
	STDMETHOD(get_PMICustomize)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_PMICustomize)(/*[in]*/ VARIANT_BOOL newVal);
	void GetMassageBoxCaption(CString& csCaption);
	STDMETHOD(get_ShowPopupMenu)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_ShowPopupMenu)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_DTE)(/*[out, retval]*/ LPDISPATCH *pVal);
	STDMETHOD(put_DTE)(/*[in]*/ LPDISPATCH newVal);
	STDMETHOD(get_ActiveTab)(/*[out, retval]*/ IToolBoxTab* *pVal);
	STDMETHOD(get_ToolBoxTabs)(/*[out, retval]*/ IToolBoxTabs* *pVal);

	BOOL m_bEnabled;
	CComPtr<IFontDisp> m_pFont;
	BOOL m_bTabStop;

	BOOL ShowTab(int nTabID, BOOL bShow);
	CToolBarCtrl& GetToolBarCtrl(int nBarID);
	
	void MaximizeTab(int nTabID);
	BOOL DeleteTab(int nTabID);
	int InsertTab(int nIndex, BSTR bstrName);
	BOOL PreTranslateAccelerator(LPMSG pMsg, HRESULT& hRet);
	
protected:
	HRESULT OnDraw(ATL_DRAWINFO& di);
	LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnLButtonDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		bHandled=TRUE;
		return 0;
	}
	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnContextMenu(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	LRESULT OnCharEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnCtlColorEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnGetDlgCodeEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnGetDlgCode(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	LRESULT OnRebarLayoutChanged(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnRelayToolTip(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	
protected:
	CCommandManager m_cmdManager;
	void _setupFonts();
	void _setupTabPopupMenu(CMenuHandle& menu, int nTabIndex, IToolBoxTab* pTab);
	void _setupItemPopupMenu(CMenuHandle& menu, CToolBoxTab* pTab, int nItemIdx, CToolBarCtrl& tb);

	CComObject<CToolBoxTabs>* m_pTabs;
	CMDICommandBarXPCtrl m_CmdBar;
	CAccessBarCtrl m_wndRebar;
	CContainedWindow m_wndEdit;
	CComBool m_bShowPopupMenu;
	CComBool m_bPMICustomize;

	CComDispatchDriver m_DTEDriver;

	CAccelerator m_accel;
	
};

#endif //__TOOLBOX_H_
