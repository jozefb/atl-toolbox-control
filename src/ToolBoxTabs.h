// ToolBoxTabs.h : Declaration of the CToolBoxTabs

#ifndef __TOOLBOXTABS_H_
#define __TOOLBOXTABS_H_

#include "resource.h"       // main symbols
#include <vector>
#include "ToolBoxTab.h"

class CToolbox;
/////////////////////////////////////////////////////////////////////////////
// CToolBoxTabs
class ATL_NO_VTABLE CToolBoxTabs : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CToolBoxTabs, &CLSID_ToolBoxTabs>,
	public ISupportErrorInfo,
	public IDispatchImpl<IToolBoxTabs, &IID_IToolBoxTabs, &LIBID_IFSTOOLBOXLib>
{
public:
	CToolBoxTabs() : m_pToolbox(NULL)
	{
		
	}

DECLARE_NO_REGISTRY()

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CToolBoxTabs)
	COM_INTERFACE_ENTRY(IToolBoxTabs)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IToolBoxTabs
public:
	bool ContainsHiddenItems();
	STDMETHOD(Add)(/*[in]*/ BSTR bstrName, /*[out, retval]*/ IToolBoxTab** pToolBoxTab);
	STDMETHOD(get_Count)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_Item)(VARIANT index, /*[out, retval]*/ IToolBoxTab* *pVal);
	STDMETHOD(get_DTE)(/*[out, retval]*/ LPDISPATCH *pVal);
	STDMETHOD(get_Parent)(/*[out, retval]*/ IToolbox* *pVal);
	STDMETHOD(get__NewEnum)(/*[out, retval]*/ LPUNKNOWN *pVal);

	STDMETHOD(FinalConstruct)();
	STDMETHOD(FinalRelease)();

	STDMETHOD(put_Parent)(CToolbox *pVal);
	HRESULT GetTabByID(int nTabID, IToolBoxTab** pTab);
	CComObject<CToolBoxTab>* GetTabByID(int nTabID);	
	HRESULT GetItemByID(int nItemID, IToolBoxItem** pItem);
	HRESULT GetItemByID(int nTabID, int nItemID, IToolBoxItem** pItem);
	void ShowAllTabs();
	HRESULT ShowTab(int nID, BSTR Name, bool bShow);
	bool TabWithNameExist(BSTR bstrName);
	CToolbox* GetToolbox();
	HRESULT ActivateTab(int nTabID);
	HRESULT RemoveTab(int nTabID);
protected:
	typedef class _tTabHolder {
		inline void _assisgn(CComObject<CToolBoxTab>* pTab)
		{
			ATLASSERT(pTab);			
			this->pTab=pTab;
			ADDREF(pTab);
		}
	public:
		_tTabHolder(CComObject<CToolBoxTab>* pTab)
		{
			_assisgn(pTab);
		}
		
		_tTabHolder() : pTab(NULL)
		{
		}

		_tTabHolder(const _tTabHolder& h)
		{
			_assisgn(h.pTab);
		}
		~_tTabHolder()
		{
			RELEASE(pTab);
		}
		CComObject<CToolBoxTab>* pTab;
		//int nID;

		_tTabHolder& operator=(const _tTabHolder& h)
		{
			_assisgn(h.pTab);			
			return *this;			
		}
		
private:

	} _tTabHolder;
	
	CToolbox* m_pToolbox;
	typedef std::vector< _tTabHolder > VectorTabs;
	VectorTabs m_vecTabs;
	typedef VectorTabs::iterator vecIter;

	static void _ReleaseTab(_tTabHolder& holder) 
	{
		holder.pTab->Release();
	};

	static void _ShowTab(_tTabHolder& holder) 
	{
		holder.pTab->put_Hidden(VARIANT_FALSE);
	};

	struct findTabByName 
	{
		BSTR bstrName;
	
		findTabByName(BSTR bstrName)
		{
			this->bstrName=bstrName;
		}
		
		bool operator()(_tTabHolder& holder)
		{
			CComBSTR tmp;
			holder.pTab->get_Name(&tmp);
			return tmp == bstrName;
		}
		
	};

	struct findHiddenTab 
	{
		
		bool operator()(_tTabHolder& holder)
		{
			CComBool bHidden;
			holder.pTab->get_Hidden(&bHidden);
			return bHidden;
		}
		
	};

	struct findTabByID 
	{
		int nID;
	
		findTabByID(int nID)
		{
			this->nID=nID;
		}
		
		bool operator()(_tTabHolder& holder)
		{
			int tmp;
			holder.pTab->get_ID(&tmp);
			return nID == tmp;
		}
		
	};

	class _CopyIDispatch
	{
	public:
		static HRESULT copy(VARIANT* pVar, _tTabHolder* ppItem) 
		{
			if(!ppItem || !pVar || !ppItem->pTab){ return E_POINTER; }
			
			::VariantInit(pVar);
			pVar->vt = VT_DISPATCH;
			return  ppItem->pTab->QueryInterface(IID_IDispatch, (void**)&pVar->pdispVal);
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
};

#endif //__TOOLBOXTABS_H_
