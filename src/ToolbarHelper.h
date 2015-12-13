//

#ifndef _TOOLBARHELPER_H_
#define _TOOLBARHELPER_H_

#pragma once

HWND CreateEmptyToolBarCtrl(HWND hWndParent, UINT nResourceID, BOOL bInitialSeparator = FALSE, 
							DWORD dwStyle = ATL_SIMPLE_TOOLBAR_STYLE, UINT nID = ATL_IDW_TOOLBAR);
#endif