// ToolBoxItem.h : Declaration of the CToolBoxItem

#ifndef __TOOLBOXITEM_H_
#define __TOOLBOXITEM_H_

#include "resource.h"       // main symbols
#include "PersistHelper.h"

class CToolBoxItems;
/////////////////////////////////////////////////////////////////////////////
// CToolBoxItem
class ATL_NO_VTABLE CToolBoxItem : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CToolBoxItem, &CLSID_ToolBoxItem>,
	public ISupportErrorInfo,
	public IDispatchImpl<IToolBoxItem, &IID_IToolBoxItem, &LIBID_IFSTOOLBOXLib>,
	public IPersistStorageImpl<CToolBoxItem>,
	public IPersistPropertyBagImpl<CToolBoxItem>,
	public IPersistStreamInitImpl<CToolBoxItem>,
	public CPersistHelper<CToolBoxItem>
{
public:
	CToolBoxItem();

DECLARE_REGISTRY_RESOURCEID(IDR_TOOLBOXITEM)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CToolBoxItem)
	COM_INTERFACE_ENTRY(IToolBoxItem)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(IPersistPropertyBag)
END_COM_MAP()

BEGIN_PROP_MAP(CToolBoxItem)
	PROP_ENTRY("Name", 1, CLSID_NULL)
	PROP_ENTRY("Guid", 2, CLSID_NULL)
	PROP_ENTRY("Format", 7, CLSID_NULL)
	PROP_ENTRY("Description", 9, CLSID_NULL)
	PROP_ENTRY("Hidden", 10, CLSID_NULL)
	PROP_ENTRY("Deleteable", 11, CLSID_NULL)
	PROP_ENTRY("Data", 8, CLSID_NULL)
// Example entries
// PROP_ENTRY("Property Description", dispid, clsid)
// PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IPersistStreamInit
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty);
	STDMETHOD(Load)(LPSTREAM pStm);

	STDMETHOD(FinalConstruct)();
	STDMETHOD(FinalRelease)();
// IToolBoxItem
public:
	STDMETHOD(get_Deleteable)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_Deleteable)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_Hidden)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_Hidden)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_Description)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Description)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_Data)(/*[out, retval]*/ VARIANT *pVal);
	STDMETHOD(put_Data)(/*[in]*/ VARIANT newVal);
	STDMETHOD(get_Format)(/*[out, retval]*/ ToolBoxItemFormat *pVal);
	STDMETHOD(put_Format)(/*[in]*/ ToolBoxItemFormat newVal);
	STDMETHOD(get_DTE)(/*[out, retval]*/ LPDISPATCH *pVal);
	STDMETHOD(get_ID)(/*[out, retval]*/ int *pVal);
	STDMETHOD(put_ID)(/*[in]*/ int newVal);
	STDMETHOD(get_Collection)(/*[out, retval]*/ IToolBoxItems* *pVal);
	STDMETHOD(Select)();
	STDMETHOD(Delete)();
	STDMETHOD(get_Guid)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Guid)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Name)(/*[in]*/ BSTR newVal);

	STDMETHOD(put_Collection)(CToolBoxItems *pVal);
		
	bool IsSelected();		
protected:
	CComBSTR m_bstrName;
	CComBSTR m_bstrGuid;
	CComBSTR m_bstrDescription;
	CToolBoxItems* m_pCollection;
	int m_nID;
	ToolBoxItemFormat m_eFormat;
	CComVariant m_varData;
	CComBool m_bHidden;
	CComBool m_bDeleteable;
};

#endif //__TOOLBOXITEM_H_
