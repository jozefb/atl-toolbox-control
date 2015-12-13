// ToolBoxTab.h : Declaration of the CToolBoxTab

#ifndef __TOOLBOXTAB_H_
#define __TOOLBOXTAB_H_

#include "resource.h"       // main symbols
#include "PersistHelper.h"
#include "IfsToolboxCP.h"

class CToolBoxTabs;
class CToolBoxItems;
class CToolbox;
/////////////////////////////////////////////////////////////////////////////
// CToolBoxTab
class ATL_NO_VTABLE CToolBoxTab : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CToolBoxTab, &CLSID_ToolBoxTab>,
	public ISupportErrorInfo,
	public IDispatchImpl<IToolBoxTab, &IID_IToolBoxTab, &LIBID_IFSTOOLBOXLib>,
	public IPersistStorageImpl<CToolBoxTab>,
	public IPersistPropertyBagImpl<CToolBoxTab>,
	public IPersistStreamInitImpl<CToolBoxTab>,
	public CPersistHelper<CToolBoxTab>
{
public:
	CToolBoxTab();

DECLARE_REGISTRY_RESOURCEID(IDR_TOOLBOXTAB)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CToolBoxTab)
	COM_INTERFACE_ENTRY(IToolBoxTab)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(IPersistPropertyBag)
END_COM_MAP()

BEGIN_PROP_MAP(CToolBoxTab)
	PROP_ENTRY("Name", 1, CLSID_NULL)
	PROP_ENTRY("Guid", 2, CLSID_NULL)
	PROP_ENTRY("ListView", 6, CLSID_NULL)
	PROP_ENTRY("Deleteable", 9, CLSID_NULL)
	PROP_ENTRY("Hidden", 10, CLSID_NULL)
	//PROP_ENTRY("ID", 7, CLSID_NULL)
	// Example entries
	// PROP_ENTRY("Property Description", dispid, clsid)
	// PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IPersistStreamInit
	
	STDMETHOD(FinalConstruct)();
	STDMETHOD(FinalRelease)();
// IToolBoxTab
public:
	
	STDMETHOD(get_Hidden)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_Hidden)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_DTE)(/*[out, retval]*/ LPDISPATCH *pVal);
	STDMETHOD(get_ID)(/*[out, retval]*/ int *pVal);
	STDMETHOD(put_ID)(/*[in]*/ int newVal);
	STDMETHOD(get_ListView)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_ListView)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_Collection)(/*[out, retval]*/ IToolBoxTabs* *pVal);
	STDMETHOD(Activate)();
	STDMETHOD(Delete)();
	STDMETHOD(get_Guid)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Guid)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Name)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_Deleteable)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_Deleteable)(/*[in]*/ VARIANT_BOOL newVal);

	HRESULT GetSelectedItem(IToolBoxItem** ppItem);
	CToolBoxItems* GetItems();
	CToolbox* GetToolbox();
	CToolBarCtrl& GetToolBarCtrl();
	CToolBoxTabs* get_Tabs(){ return m_pTabs; };
	STDMETHOD(get_ToolBoxItems)(/*[out, retval]*/ IToolBoxItems* *pVal);
	
	void put_Collection(CToolBoxTabs* pTabs);
	HRESULT GetItem(int nItemID, IToolBoxItem** ppItem);

protected:
	CComBSTR m_bstrName;
	CComBSTR m_bstrGuid;
	CComBool m_bListView;
	CComBool m_bDeleteable;
	CComBool m_bHidden;
	CToolBoxTabs* m_pTabs;
	int m_nID;
	CComObject<CToolBoxItems>* m_pItems;
};

#endif //__TOOLBOXTAB_H_
