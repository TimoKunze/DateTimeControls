// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.calU) {
		DispEventUnadvise(controls.calU);
		controls.calU.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);

	calUWnd.SubclassWindow(GetDlgItem(IDC_CALU));
	calUWnd.QueryControl(__uuidof(ICalendar), reinterpret_cast<LPVOID*>(&controls.calU));
	if(controls.calU) {
		DispEventAdvise(controls.calU);
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.aboutButton.IsWindow() && controls.calU) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			CRect aboutRectangle;
			controls.aboutButton.GetWindowRect(&aboutRectangle);

			OLE_XPOS_PIXELS left = 0;
			OLE_XPOS_PIXELS right = clientRectangle.Width() - 16;
			OLE_YPOS_PIXELS top = 0;
			OLE_YPOS_PIXELS bottom = clientRectangle.Height() - 8 - aboutRectangle.Height() - 16;
			controls.calU->MinimizeRectangle(&left, &top, &right, &bottom);
			calUWnd.MoveWindow((clientRectangle.Width() - (right - left)) / 2, 8, right - left, bottom - top);
			controls.aboutButton.MoveWindow((clientRectangle.Width() - aboutRectangle.Width()) / 2, clientRectangle.Height() - 8 - aboutRectangle.Height(), aboutRectangle.Width(), aboutRectangle.Height());
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnWindowPosChanging(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.aboutButton.IsWindow() && controls.calU) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			OLE_XSIZE_PIXELS minimumMCWidth = 0;
			OLE_YSIZE_PIXELS minimumMCHeight = 0;
			controls.calU->GetMinimumSize(&minimumMCWidth, &minimumMCHeight);
			if(pDetails->cx < minimumMCWidth + 50) {
				pDetails->cx = minimumMCWidth + 50;
			}
			CRect aboutRectangle;
			controls.aboutButton.GetWindowRect(&aboutRectangle);
			if(pDetails->cy < minimumMCHeight + aboutRectangle.Height() + 70) {
				pDetails->cy = minimumMCHeight + aboutRectangle.Height() + 70;
			}
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.calU) {
		controls.calU->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}
