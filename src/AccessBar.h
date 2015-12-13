#if !defined(AFX_ACCESSBAR_H__20020628_366A_EF2D_9C7E_0080AD509054__INCLUDED_)
#define AFX_ACCESSBAR_H__20020628_366A_EF2D_9C7E_0080AD509054__INCLUDED_

#pragma once

/////////////////////////////////////////////////////////////////////////////
// CAccessBarCtrl - MS Access Navigation Bar
//
// Written by Bjarke Viksoe (bjarke@viksoe.dk)
// Copyright (c) 2002 Bjarke Viksoe.
//
// This code may be used in compiled form in any way you desire. This
// file may be redistributed by any means PROVIDING it is 
// not sold for profit without the authors written consent, and 
// providing that this notice and the authors name is included. 
//
// This file is provided "as is" with no expressed or implied warranty.
// The author accepts no liability if it causes any damage to you or your
// computer whatsoever. It's free, so don't hassle me about it.
//
// Beware of bugs.
//

#ifndef __cplusplus
  #error ATL requires C++ compilation (use a .cpp suffix)
#endif

#ifndef __ATLAPP_H__
  #error This control requires atlapp.h to be included first
#endif

#ifndef __ATLCTRLS_H__
  #error This control requires atlctrls.h to be included first
#endif

#if (_WTL_VER < 0x0700)
   #error This control requires WTL version 7.0 or higher
#endif

#if (_WIN32_IE < 0x0400)
  #error This control requires _WIN32_IE >= 0x0400
#endif

// From Platform SDK for MSIE >= 0x0501
#ifndef BTNS_SHOWTEXT
#define BTNS_SHOWTEXT           0x0040
#endif  // BTNS_SHOWTEXT

typedef CWinTraitsOR< RBS_VERTICALGRIPPER | RBS_VARHEIGHT  | CCS_VERT|CCS_NODIVIDER | CCS_NORESIZE | CCS_NOPARENTALIGN, 0 > CAccessBarCtrlTraits;

template< class T, class TBase = CReBarCtrl, class TWinTraits = CAccessBarCtrlTraits >
class ATL_NO_VTABLE CAccessBarImpl : 
   public CWindowImpl< T, TBase, TWinTraits >,
   public CCustomDraw< T >
{
public:
   DECLARE_WND_SUPERCLASS(NULL, TBase::GetWndClassName())

   enum 
   {
      _MaxButtonTextLength = 128,
   };

	typedef struct tagACCESSBAR
	{
		CToolBarCtrl ctrlToolBar;
		CPagerCtrl ctrlPager;
	} ACCESSBAR;

   // Operations

   BOOL SubclassWindow(HWND hWnd)
   {
      ATLASSERT(m_hWnd == NULL);
      ATLASSERT(::IsWindow(hWnd));
#ifdef _DEBUG
      TCHAR szBuffer[32] = { 0 };
      if( ::GetClassName(hWnd, szBuffer, (sizeof(szBuffer)/sizeof(TCHAR))-1) ) {
         ATLASSERT(::lstrcmpi(szBuffer, TBase::GetWndClassName())==0);
      }
#endif
      BOOL bRet = CWindowImpl< T, TBase, TWinTraits >::SubclassWindow(hWnd);
      if( bRet ) _Init();
      return bRet;
   }

   void SetHighlightedBandID(int nBandID)
   {
	   m_nHighlightedBandID = nBandID;
	   Invalidate();
	   RedrawWindow();
   }

   int InsertBand(int iIndex, LPCTSTR pstrTitle, UINT nRes)
   {
	   ATLASSERT(iIndex>=-1);
	   CRect rcClient;
	   GetClientRect(&rcClient);
      ATLASSERT(!::IsBadStringPtr(pstrTitle, -1));
      ATLASSERT(nRes!=0);
      // Create the child controls
      // First we create the Pager control to allow scrolling of panes
      ACCESSBAR Band;
      DWORD dwStyle;
      dwStyle = WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | PGS_VERT;
      Band.ctrlPager.Create(m_hWnd, &rcDefault, NULL, dwStyle);
	  // Resize the pager
	  Band.ctrlPager.ResizeClient(rcClient.Width(), -1);
      ATLASSERT(Band.ctrlPager.IsWindow()); // Remember to enable ICC_PAGESCROLLER_CLASS in ::InitCommonControlsEx() call!
      // Then we create the ToolBar control, which acts as the view itself
      dwStyle = WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | 
		  CCS_NODIVIDER | CCS_LEFT | CCS_NORESIZE | CCS_NOPARENTALIGN | TBSTYLE_FLAT | TBSTYLE_LIST|TBSTYLE_TOOLTIPS;
	  if( nRes == -1 )
	  {
		Band.ctrlToolBar = CreateEmptyToolBarCtrl(Band.ctrlPager, nRes, FALSE, dwStyle);
	  }
	  else
	  {
		Band.ctrlToolBar = CFrameWindowImplBase<>::CreateSimpleToolBarCtrl(Band.ctrlPager, nRes, FALSE, dwStyle);
	  }
      ATLASSERT(Band.ctrlToolBar.IsWindow());
      // Well done! Now, add the ToolBar to the Pager control and prepare ToolBar buttons...
      _PrepareToolBar(Band);
      _PreparePager(Band);
      
      // Finally we can add the ReBar band...
	  int nCount = m_Bands.GetSize();
      REBARBANDINFO rbi = { 0 };
      rbi.cbSize = sizeof(REBARBANDINFO);  
      rbi.fMask = RBBIM_TEXT | RBBIM_STYLE | RBBIM_ID|RBBIM_CHILD | RBBIM_SIZE | RBBIM_CHILDSIZE | RBBIM_LPARAM|RBBIM_HEADERSIZE;
      rbi.lpText = const_cast<LPTSTR>(pstrTitle);
      rbi.cx = rcClient.Width(); 
      rbi.cxMinChild = 0;
	  rbi.lParam = nCount; // Real index of band data in array;
      rbi.cyMinChild = rcClient.Width();
      rbi.fStyle = RBBS_GRIPPERALWAYS/*|RBBS_CHILDEDGE*/;
      rbi.hwndChild  = Band.ctrlPager;
	  rbi.wID = ++m_nNextBandID;
	  rbi.cxHeader = m_nBandHeaderHeight;
      BOOL bOk = TBase::InsertBand(iIndex, &rbi);
	  m_Bands.Add(Band);
      if( bOk ) {
         // First pane should always be initially maximized
		  MaximizeBand(iIndex!=-1?iIndex:rbi.wID);
		 m_nMaximizedBandID=rbi.wID;
      }
	  
      return rbi.wID;
   }

	void MaximizeBand(int nBandIndex, BOOL bIdeal = FALSE)
	{
		REBARBANDINFO rbi={0};
		rbi.cbSize = sizeof(rbi);
		rbi.fMask=RBBIM_ID;
		GetBandInfo(nBandIndex, &rbi);
		m_nMaximizedBandID=rbi.wID;
		TBase::MaximizeBand(nBandIndex, bIdeal);
	}


   int GetMaximizedBandID()
   {
	   return m_nMaximizedBandID;
   }

   int GetBarHeaderHeight()
   {
	   return m_nBandHeaderHeight;
   }

	void SetFont(IFontDisp* pFontDisp)
	{
		CComQIPtr<IFont> pFont = pFontDisp;
		if( pFont )
		{
			HFONT hFont=NULL;
			pFont->get_hFont(&hFont);
			TBase::SetFont(hFont, FALSE);
			CY c;
			pFont->get_Size(&c);
			m_nBandHeaderHeight = 2*(c.Lo/10000);
			if( m_nBandHeaderHeight < 18 )
			{
				m_nBandHeaderHeight = 18;
			}
		}
	}

   BOOL SetBandName(int nBandIdx, LPCTSTR lpszName)
   {
	   REBARBANDINFO rbi={0};
	   rbi.cbSize = sizeof(rbi);
	   rbi.fMask=RBBIM_TEXT;
	   rbi.lpText = (LPSTR)lpszName;
	   rbi.cch=_MaxButtonTextLength;
	   return SetBandInfo(nBandIdx, &rbi);
   }
   CToolBarCtrl& GetToolBarCtrl(int iBandIdx) const
   {
      ATLASSERT(iBandIdx>=0 && iBandIdx<m_Bands.GetSize());
	  REBARBANDINFO rbi={0};
	  rbi.cbSize = sizeof(rbi);
	  rbi.fMask=RBBIM_LPARAM;
	  ATLVERIFY(GetBandInfo(iBandIdx, &rbi));
	  ATLASSERT(rbi.lParam>=0 && rbi.lParam<m_Bands.GetSize());
      return m_Bands[rbi.lParam].ctrlToolBar;
   }

   ACCESSBAR& GetBandData(int iBandIdx)
   {
	   ATLASSERT(iBandIdx>=0 && iBandIdx<m_Bands.GetSize());
	   REBARBANDINFO rbi={0};
	   rbi.cbSize = sizeof(rbi);
	   rbi.fMask=RBBIM_LPARAM;
	   ATLVERIFY(GetBandInfo(iBandIdx, &rbi));
	   ATLASSERT(rbi.lParam>=0 && rbi.lParam<m_Bands.GetSize());
	   return m_Bands[rbi.lParam];
   }



   // Implementation

   BOOL _Init()
   {
	   m_nNextBandID = -1;
	   m_nHighlightedBandID=-1;
	   m_nBandHeaderHeight = 19;
	   m_nMaximizedBandID =-1;
      ATLASSERT(::IsWindow(m_hWnd));
      ATLASSERT(GetDlgCtrlID()!=ATL_IDW_TOOLBAR);
      // Initialize and send the REBARINFO structure.
      REBARINFO rbi = { 0 };
      rbi.cbSize = sizeof(REBARINFO);
      rbi.fMask = 0;
      rbi.himl = NULL;
      SetBarInfo(&rbi);

	  m_blackBrush.CreateSolidBrush(RGB(0,0,0));
	  m_clrBtnText = m_clrRebarText = ::GetSysColor(COLOR_BTNTEXT);
	  m_clrBtnTextHighlightt = ::GetSysColor(COLOR_WINDOW);
	  m_clrRebarBkg = ::GetSysColor(COLOR_BTNFACE);

      // Assign rebar colors
      SetTextColor(m_clrRebarText);
      SetBkColor(m_clrRebarBkg);
	  return TRUE;
   }

   void _PrepareToolBar(ACCESSBAR& bar) const
   {
		CRect rcClient;
		GetClientRect(&rcClient);
      CToolBarCtrl& tb = bar.ctrlToolBar;
      int nCount = tb.GetButtonCount();
      for( int i=0; i<nCount; i++ ) {
         // Get button info from index - so we can determine the Command Identifier
         TBBUTTON btn = { 0 };
         tb.GetButton(i, &btn);
         // Read existing button style
         TBBUTTONINFO tbi = { 0 };
         tbi.cbSize = sizeof(TBBUTTONINFO),
         tbi.dwMask = TBIF_STYLE;
         tb.GetButtonInfo(btn.idCommand, &tbi);
         // Change button style and text
         tbi.dwMask = TBIF_STYLE | TBIF_TEXT | TBIF_SIZE;
         tbi.fsStyle |= TBSTYLE_BUTTON | BTNS_SHOWTEXT | TBSTYLE_CHECK;
         tbi.cx = rcClient.Width();
         // Load button text (use part after newline)...
         TCHAR szText[_MaxButtonTextLength+1] = { 0 };
         ::LoadString(_Module.GetResourceInstance(), btn.idCommand, szText, _MaxButtonTextLength);
         LPTSTR pStr = szText;
         while( *pStr && *pStr!=_T('\n') ) pStr = ::CharNext(pStr);
         ATLASSERT(*pStr); // Must contain a newline char; like this "Bla bla bla\nCommand" !!!
         *pStr = _T(' ');
         tbi.pszText = pStr;
         // Assign new button style
         tb.SetButtonInfo(btn.idCommand, &tbi);
      }
      // Set the row count to a fixed number (ensures only 1 button pr line)
      RECT rc;
      tb.SetRows(nCount, FALSE, &rc);
      // Change the size of the buttons to allow some padding between the text and edges
      SIZE size;
      tb.GetButtonSize(size);
      size.cy += 4;
      tb.SetButtonSize(size);
      // Change the button padding as well
      tb.GetPadding(&size),
      size.cx += 2;
      size.cy += 4;
      tb.SetPadding(size.cx, size.cy);
      // I would like these text styles...
      tb.SetDrawTextFlags(DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_CENTER | DT_BOTTOM, 
                          DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
      // and these border colours when the mouse hovers over the toolbutton
      COLORSCHEME color;
      color.dwSize = sizeof(COLORSCHEME);
      color.clrBtnShadow = ::GetSysColor(COLOR_3DDKSHADOW);
      color.clrBtnHighlight = ::GetSysColor(COLOR_BTNHIGHLIGHT);
      tb.SetColorScheme(&color);
   }
   void _PreparePager(ACCESSBAR& bar)
   {
      bar.ctrlPager.SetBkColor(GetBkColor());
      bar.ctrlPager.SetChild(bar.ctrlToolBar);
   }

   // Message map and handler

   BEGIN_MSG_MAP(CAccessBarImpl)
      MESSAGE_HANDLER(WM_CREATE, OnCreate)
      MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
      MESSAGE_HANDLER(WM_SIZE, OnSize)
	  MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
      NOTIFY_CODE_HANDLER(PGN_CALCSIZE, OnCalcSize)
      REFLECTED_NOTIFY_CODE_HANDLER(RBN_BEGINDRAG, OnRebarBeginDrag)
	  REFLECTED_NOTIFY_CODE_HANDLER(RBN_ENDDRAG, OnRebarEndDrag)
      COMMAND_RANGE_HANDLER(0x0000, 0xFFFF, OnCommand)  // Trap all command notifications
      CHAIN_MSG_MAP_ALT( CCustomDraw< T >, 1 )          // Custom draw of rebar
      CHAIN_MSG_MAP( CCustomDraw< T > )                 // Custom draw of toolbar
      DEFAULT_REFLECTION_HANDLER()
   END_MSG_MAP()
 
   LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
   {
      LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
      if( !_Init() )
	  {
		  return -1;
	  }
      return lRes;
   }
   
   LRESULT OnLButtonDown(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
   {
	   POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	   RBHITTESTINFO rbhi;
	   rbhi.pt=pt;
	   HitTest(&rbhi);
	   REBARBANDINFO rbi={0};
	   rbi.cbSize = sizeof(rbi);
	   rbi.fMask=RBBIM_ID;
	   GetBandInfo(rbhi.iBand, &rbi);
	   m_nMaximizedBandID=rbi.wID;
	   bHandled=FALSE;
	   //::SendMessage(GetParent(), WM_LBUTTONDOWN, wParam, lParam);
	   return 0;
   }
   LRESULT OnSetCursor(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
   {
      // HACK: Prevent drag-cursor when clicking on header
      ::SetCursor(::LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)));
      return 0;
   }
   LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
   {
      LRESULT lRes = DefWindowProc();
      // Hmm, resizing has proven quite tedious. The rebar and toolbar controls
      // prefer to stick to a container side and takes any oppotunity to resize itself.
      // I guess figuring out the right combination of CCS_xxx flags would solve the problem
      // but for now I'll just force a resize on child controls/buttons...
      CRect rcClient;
      GetClientRect(&rcClient);
      int cx = rcClient.Width();
	  TBBUTTONINFO tbi = { 0 };
	  tbi.cbSize = sizeof(TBBUTTONINFO),
	  tbi.dwMask = TBIF_SIZE;
      for( int i=0; i<m_Bands.GetSize(); i++ ) {
         // Reisze the rebar band
         REBARBANDINFO rbi = { 0 };
         rbi.cbSize = sizeof(REBARBANDINFO);  
         rbi.fMask = RBBIM_CHILDSIZE;
         rbi.cxMinChild = 0;
         rbi.cyMinChild = cx;
         SetBandInfo(i, &rbi);
         // Resize the pager
         ACCESSBAR& data = m_Bands[i];
		 data.ctrlPager.ResizeClient(cx, -1);
         // Resize tool buttons
         CToolBarCtrl& tb = data.ctrlToolBar;
		 BOOL bIsList = tb.GetStyle() & TBSTYLE_LIST;
         int nButtons = tb.GetButtonCount();
         for( int j=0; j<nButtons && bIsList; j++ ) {
            TBBUTTON btn = { 0 };
            tb.GetButton(j, &btn);
            tbi.cx = cx;
            tb.SetButtonInfo(btn.idCommand, &tbi);
         }
      }
      return lRes;
   }
   LRESULT OnCommand(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
   {
      // Remove checked state for other tool buttons
      for( int i=0; i<m_Bands.GetSize(); i++ ) {
         CToolBarCtrl& tb = m_Bands[i].ctrlToolBar;
         int nButtons = tb.GetButtonCount();
         for( int j=0; j<nButtons; j++ ) {
            TBBUTTON btn = { 0 };
            tb.GetButton(j, &btn);
            if( btn.idCommand != wID ) 
			{
				tb.CheckButton(btn.idCommand, FALSE);
			}

            if( btn.idCommand == wID && !tb.IsButtonChecked(btn.idCommand) )
			{
				tb.CheckButton(btn.idCommand, TRUE);
			}
         }
      }
      // We've overwritten the WM_COMMAND from a Pager control, and
      // we'll forward it to the top-level window.
      return ::SendMessage(GetParent(), WM_COMMAND, MAKEWPARAM(wID, 0), (LPARAM) m_hWnd);
   }
   LRESULT OnCalcSize(int /*idCtrl*/, LPNMHDR pnmh, BOOL& /*bHandled*/)
   {
      // Get the optimum height of the Toolbar for the Pager.
      LPNMPGCALCSIZE pCalcSize = (LPNMPGCALCSIZE) pnmh;
      switch( pCalcSize->dwFlag ) {
      case PGF_CALCHEIGHT:
         {
            CToolBarCtrl tb = ::GetWindow(pnmh->hwndFrom, GW_CHILD);
            SIZE size;
            tb.GetMaxSize(&size);
            pCalcSize->iHeight = size.cy-1;
         }
         break;
      }
      return 0;
   }
   LRESULT OnRebarBeginDrag(int idCtrl, LPNMHDR pnmh, BOOL& /*bHandled*/)
   {
		LPNMREBAR pRebar = (LPNMREBAR)pnmh;
		SetHighlightedBandID(pRebar->wID);
		::SetCursor(::LoadCursor(NULL, MAKEINTRESOURCE(IDC_SIZENS)));
		return 0; // Don't allow drag'n'drop of rebar bands...
   }

   LRESULT OnRebarEndDrag(int idCtrl, LPNMHDR pnmh, BOOL& /*bHandled*/)
   {
	   LPNMREBAR pRebar = (LPNMREBAR)pnmh;
	   SetHighlightedBandID(-1);
	   ::SetCursor(::LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)));
	   return 0; // Don't allow drag'n'drop of rebar bands...
   }

   // Custom draw methods

   DWORD OnPrePaint(int idCtrl, LPNMCUSTOMDRAW /*lpNMCustomDraw*/)
   {
      // NOTE: Both the ReBar and ToolBar custom painting end up in these
      //       methods, so we need to carefully separate them...
      return CDRF_NOTIFYITEMDRAW;
   }
   DWORD OnItemPrePaint(int idCtrl, LPNMCUSTOMDRAW lpNMCustomDraw)
   {
      // If this is a toolbar custom drawing, then change colours
      if( idCtrl == ATL_IDW_TOOLBAR ) {
		 CDCHandle dc( lpNMCustomDraw->hdc );
		 HFONT hOldFont = dc.SelectFont(GetFont());
         LPNMTBCUSTOMDRAW lpNMTBCD = (LPNMTBCUSTOMDRAW) lpNMCustomDraw;
       //  lpNMTBCD->hbrMonoDither = ::GetSysColorBrush(COLOR_BTNHIGHLIGHT);
       //  lpNMTBCD->clrText = m_clrBtnText;
       //  lpNMTBCD->clrTextHighlight = m_clrBtnTextHighlightt;
         return CDRF_NEWFONT;
      }
      // Otherwise continue to custom paint the rebar grippers...
      return CDRF_NOTIFYPOSTPAINT;
   }
   DWORD OnItemPostPaint(int /*idCtrl*/, LPNMCUSTOMDRAW lpNMCustomDraw)
   {
		// Custom paint the gripper bars in the rebar bands.

		CDCHandle dc( lpNMCustomDraw->hdc );
		// Paint a nice header-button instead.
		//dc.DrawFrameControl(&rcItem, DFC_BUTTON, DFCS_BUTTONPUSH);
		int nIndex=IdToIndex(lpNMCustomDraw->dwItemSpec);
		REBARBANDINFO rbi={0};
		rbi.cbSize = sizeof(rbi);
		rbi.fMask=RBBIM_LPARAM;
		ATLVERIFY(GetBandInfo(nIndex, &rbi));
		ATLASSERT(rbi.lParam>=0 && rbi.lParam<m_Bands.GetSize());
		CRect rcFill=lpNMCustomDraw->rc;
		rcFill.DeflateRect(1,1,1,1);
		dc.FillRect(&rcFill, (m_nHighlightedBandID == lpNMCustomDraw->dwItemSpec) ? m_blackBrush : ::GetSysColorBrush(COLOR_BTNFACE));
		dc.Draw3dRect(&lpNMCustomDraw->rc, ::GetSysColor(COLOR_BTNHIGHLIGHT),::GetSysColor(COLOR_BTNSHADOW));

		// And to draw header text
		// HACK: There is a bug in the NMCUSTOMDRAW structure for rebars. Apparently the
		//       'dwItemSpec' member does not contain the gripper index - it is always 0.
		//       We use the 'lParam' instead to figure out which header it is.
		TCHAR szText[_MaxButtonTextLength+1] = { 0 };
		rbi.fMask = RBBIM_TEXT;
		rbi.lpText = szText;
		rbi.cch = _MaxButtonTextLength;
		GetBandInfo(nIndex, &rbi);

		HFONT hOldFont = dc.SelectFont(GetFont());
		dc.SetBkMode(TRANSPARENT);
		dc.SetTextColor((m_nHighlightedBandID == lpNMCustomDraw->dwItemSpec) ? RGB(255,255,255) : m_clrRebarText);
		dc.DrawText(szText, -1, &lpNMCustomDraw->rc, DT_SINGLELINE | DT_VCENTER | DT_CENTER|DT_END_ELLIPSIS);
		dc.SelectFont(hOldFont);

		return CDRF_SKIPDEFAULT;
   }

protected:
	int m_nBandHeaderHeight;

	CSimpleArray<ACCESSBAR> m_Bands;
	UINT m_clrRebarText;
	UINT m_clrRebarBkg;
	UINT m_clrBtnText;
	UINT m_clrBtnTextHighlightt;
	int m_nMaximizedBandID;
	int m_nHighlightedBandID;
	int m_nNextBandID;

	CBrush m_blackBrush;
};

class CAccessBarCtrl : public CAccessBarImpl<CAccessBarCtrl>
{
public:
   DECLARE_WND_SUPERCLASS(_T("WTL_AccessBar"), GetWndClassName())  
};




#endif // !defined(AFX_ACCESSBAR_H__20020628_366A_EF2D_9C7E_0080AD509054__INCLUDED_)


