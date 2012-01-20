// 

#include "StdAfx.h"
#include "ToolbarHelper.h"

HWND CreateEmptyToolBarCtrl(HWND hWndParent, UINT nResourceID, BOOL bInitialSeparator, 
							DWORD dwStyle, UINT nID)
{
#ifndef _WIN32_WCE
#if (_ATL_VER >= 0x0700)
	HWND hWnd = ::CreateWindowEx(0, TOOLBARCLASSNAME, NULL, dwStyle, 0,0,100,100,
		hWndParent, (HMENU)LongToHandle(nID), ATL::_AtlBaseModule.GetModuleInstance(), NULL);
#else //!(_ATL_VER >= 0x0700)
	HWND hWnd = ::CreateWindowEx(0, TOOLBARCLASSNAME, NULL, dwStyle, 0,0,100,100,
		hWndParent, (HMENU)LongToHandle(nID), _Module.GetModuleInstance(), NULL);
#endif //!(_ATL_VER >= 0x0700)
	if(hWnd == NULL)
	{
		ATLASSERT(FALSE);
		return NULL;
	}
#else // CE specific
	dwStyle;
	nID;
	// The toolbar must go onto the existing CommandBar or MenuBar
	HWND hWnd = hWndParent;
#endif //_WIN32_WCE
	
	::SendMessage(hWnd, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0L);
	
	return hWnd;
}


BOOL GetDefaultGuiFont(CComPtr<IFontDisp>& pIfontDisp)
{
	BOOL bRet= FALSE;
	CFontHandle hFont = AtlGetDefaultGuiFont();
	if( hFont )
	{
		LOGFONT lf={0};
		hFont.GetLogFont(&lf);
		
		FONTDESC FD={0};
		FD.cbSizeofstruct=sizeof(FD);
		FD.fItalic=lf.lfItalic;
		FD.fStrikethrough = lf.lfStrikeOut;
		FD.fUnderline = lf.lfUnderline;
		CComBSTR bstrName=lf.lfFaceName;
		FD.lpstrName = bstrName;
		FD.sCharset = lf.lfCharSet;
		FD.sWeight = lf.lfWeight;
		FD.cySize.Lo = lf.lfHeight*10000;

		HRESULT hr = OleCreateFontIndirect(&FD,IID_IFontDisp,(void**)&pIfontDisp);
		if( SUCCEEDED(hr))
		{
			bRet=TRUE;
		}
		
	}
	
	return bRet;
}