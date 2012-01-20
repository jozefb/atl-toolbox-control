// ToolBoxItems.h : Declaration of the CToolBoxItems

#ifndef __TOOLBOXITEMS_H_
#define __TOOLBOXITEMS_H_

#include "resource.h"       // main symbols
#include <vector>
#include "ToolBoxItem.h"

class CToolBoxTab;
class CToolbox;

/////////////////////////////////////////////////////////////////////////////
// CToolBoxItems
class ATL_NO_VTABLE CToolBoxItems : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CToolBoxItems, &CLSID_ToolBoxItems>,
	public ISupportErrorInfo,
	public IDispatchImpl<IToolBoxItems, &IID_IToolBoxItems, &LIBID_IFSTOOLBOXLib>
{
public:
	CToolBoxItems()
	{
		m_pTab=NULL;
		m_nBitmapAdded = 3;
	}
	
	enum {
		IMG_IDX_POINTER=0,
		IMG_IDX_HTML=1,
		IMG_IDX_TEXT=2,
	};		

DECLARE_NO_REGISTRY()

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CToolBoxItems)
	COM_INTERFACE_ENTRY(IToolBoxItems)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()


// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	STDMETHOD(FinalConstruct)();
	STDMETHOD(FinalRelease)();
// IToolBoxItems
public:
	STDMETHOD(get_DTE)(/*[out, retval]*/ LPDISPATCH *pVal);
	STDMETHOD(Add)(/*[in]*/ BSTR Name, /*[in]*/ VARIANT Data, /*[in, optional, defaultvalue(ToolBoxItemFormatText)]*/ ToolBoxItemFormat Format, /*[out, retval]*/ IToolBoxItem** ppItem);
	STDMETHOD(get_Parent)(/*[out, retval]*/ IToolBoxTab* *pVal);
	STDMETHOD(get_SelectedItem)(/*[out, retval]*/ IToolBoxItem* *pVal);
	STDMETHOD(get_Count)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_Item)(VARIANT index, /*[out, retval]*/ IToolBoxItem* *pVal);
	STDMETHOD(get__NewEnum)(/*[out, retval]*/ LPUNKNOWN *pVal);

	HRESULT Add(CComObject<CToolBoxItem>* pItem, BSTR Name, VARIANT Data);
	bool IsItemSelected(int nItemID);
	HRESULT GetItem(int nItemID, IToolBoxItem** ppItem);
	HRESULT GetSelectedItem(IToolBoxItem** ppItem);
	bool ItemWithNameExist(BSTR bstrName);
	CToolBoxTab* GetTab();
	HRESULT DeleteItem(UINT nItemID);
	HRESULT SelectItem(UINT nItemID);
	HRESULT HideItem(UINT nItemID,bool bHide);
	void put_Parent(CToolBoxTab* pTab);

protected:

	void AddToolBarButton(int nCmdID, int nIndex, int nBitmapIdx, const CString& csName);
	
	typedef class _tHolder {
		void _assisgn(CComObject<CToolBoxItem>* pItem)
		{
			ATLASSERT(pItem);			
			this->pItem=pItem;
			ADDREF(pItem);
		}
	public:
		_tHolder(CComObject<CToolBoxItem>* pItem)
		{
			_assisgn(pItem);
		}
		
		_tHolder(const _tHolder& h)
		{
			_assisgn(h.pItem);
		}
		~_tHolder()
		{
			RELEASE(pItem);
		}

		CComObject<CToolBoxItem>* pItem;
		int m_wID;
		
		_tHolder& operator=(const _tHolder& h)
		{
			_assisgn(h.pItem);			
			return *this;			
		}
	} _tHolder;

	class _CopyIDispatch
	{
	public:
		static HRESULT copy(VARIANT* pVar, _tHolder* ppItem) 
		{
			if(!ppItem || !pVar || !ppItem->pItem){ return E_POINTER; }
			
			::VariantInit(pVar);
			pVar->vt = VT_DISPATCH;
			return  ppItem->pItem->QueryInterface(IID_IDispatch, (void**)&pVar->pdispVal);
		}
		
		static void init(VARIANT* pVar){}
		static void destroy(VARIANT* pVar) 
		{
			ATLASSERT(pVar && pVar->pdispVal && pVar->vt == VT_DISPATCH);
			if(pVar && pVar->pdispVal && pVar->vt == VT_DISPATCH) 
			{
				RELEASE(pVar->pdispVal);
				pVar->vt=VT_EMPTY;
			}
		}
	};

	struct findItemByName 
	{
		BSTR bstrName;
		
		findItemByName(BSTR bstrName)
		{
			this->bstrName=bstrName;
		}
		
		bool operator()(_tHolder& holder)
		{
			ATLASSERT(holder.pItem);
			CComBSTR tmp;
			holder.pItem->get_Name(&tmp);
			return tmp == bstrName;
		}
		
	};

	struct findSelectedItem 
	{
		bool operator()(_tHolder& holder)
		{
			ATLASSERT(holder.pItem);
			return holder.pItem->IsSelected();
		}	
	};

	struct findItemByID 
	{
		int id;
		
		findItemByID(int id)
		{
			this->id=id;
		}
		
		bool operator()(_tHolder& holder)
		{
			ATLASSERT(holder.pItem);
			int tmp;
			holder.pItem->get_ID(&tmp);
			return id == tmp;
		}
		
	};

	static void releaseItem(_tHolder& holder) 
	{
		ATLASSERT(holder.pItem);
		holder.pItem->Release();
	};

//	CToolbox* m_pToolbox;
	CToolBoxTab* m_pTab;
	typedef std::vector< _tHolder > VectorItems;
	VectorItems m_vecItems;
	typedef VectorItems::iterator vecItemIter;
	int m_nBitmapAdded;
	
};

#endif //__TOOLBOXITEMS_H_
