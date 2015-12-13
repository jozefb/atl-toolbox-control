// CommandManager.h: interface for the CCommandManager class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_COMMANDMANAGER_H__261A8B7B_C9EE_4E1D_845B_D2AA4FE3CE1B__INCLUDED_)
#define AFX_COMMANDMANAGER_H__261A8B7B_C9EE_4E1D_845B_D2AA4FE3CE1B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CToolbox;
class CAccessBarCtrl;

/************************************************************************/
/* Helper class for command processing                                                                     */
/************************************************************************/
class CCommandManager  
{
public:
	int GetItemID();
	int GetTabID();
	void StopCmdProcessing();
	void SetTabAndItemID(int nTabID, int nItemID);
	int ProcessCmd(int cmd);
	LRESULT OnKeyDownEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
	LRESULT OnKillFocusEdit(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);	
	CCommandManager(CToolbox& wndToolbox);
	virtual ~CCommandManager();

protected:
	void OnEditPaste();
	void OnEditCut();
	void OnEditCopy();
	void OnRenameTabItem();
	void OnDeleteTabItem();
	void OnToggleListView();
	void OnHideTab();
	void OnMoveTabOrItem(bool bUp);
	void OnDeleteTab();
	void OnRenameTab();
	void OnAddNewTab();
	
	CContainedWindow& m_wndEdit;
	CAccessBarCtrl& m_wndRebar;
	CToolbox& m_wndTbx;
	int m_nActiveTabID;
	int m_nActiveItemID;
	int m_cmd;

};

#endif // !defined(AFX_COMMANDMANAGER_H__261A8B7B_C9EE_4E1D_845B_D2AA4FE3CE1B__INCLUDED_)
