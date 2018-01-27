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
		IDispEventImpl<IDC_CALU, CMainDlg, &__uuidof(DTCtlsLibU::_ICalendarEvents), &LIBID_DTCtlsLibU, 1, 5>::DispEventUnadvise(controls.calU);
		controls.calU.Release();
	}
	if(controls.pickerU) {
		IDispEventImpl<IDC_PICKERU, CMainDlg, &__uuidof(DTCtlsLibU::_IDateTimePickerEvents), &LIBID_DTCtlsLibU, 1, 5>::DispEventUnadvise(controls.pickerU);
		controls.pickerU.Release();
	}
	if(controls.calA) {
		IDispEventImpl<IDC_CALA, CMainDlg, &__uuidof(DTCtlsLibA::_ICalendarEvents), &LIBID_DTCtlsLibA, 1, 5>::DispEventUnadvise(controls.calA);
		controls.calA.Release();
	}
	if(controls.pickerA) {
		IDispEventImpl<IDC_PICKERA, CMainDlg, &__uuidof(DTCtlsLibA::_IDateTimePickerEvents), &LIBID_DTCtlsLibA, 1, 5>::DispEventUnadvise(controls.pickerA);
		controls.pickerA.Release();
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

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	aboutButton.SubclassWindow(GetDlgItem(ID_APP_ABOUT));

	calUWnd.SubclassWindow(GetDlgItem(IDC_CALU));
	calUWnd.QueryControl(__uuidof(DTCtlsLibU::ICalendar), reinterpret_cast<LPVOID*>(&controls.calU));
	if(controls.calU) {
		IDispEventImpl<IDC_CALU, CMainDlg, &__uuidof(DTCtlsLibU::_ICalendarEvents), &LIBID_DTCtlsLibU, 1, 5>::DispEventAdvise(controls.calU);
	}

	pickerUWnd.SubclassWindow(GetDlgItem(IDC_PICKERU));
	pickerUWnd.QueryControl(__uuidof(DTCtlsLibU::IDateTimePicker), reinterpret_cast<LPVOID*>(&controls.pickerU));
	if(controls.pickerU) {
		IDispEventImpl<IDC_PICKERU, CMainDlg, &__uuidof(DTCtlsLibU::_IDateTimePickerEvents), &LIBID_DTCtlsLibU, 1, 5>::DispEventAdvise(controls.pickerU);
	}

	calAWnd.SubclassWindow(GetDlgItem(IDC_CALA));
	calAWnd.QueryControl(__uuidof(DTCtlsLibA::ICalendar), reinterpret_cast<LPVOID*>(&controls.calA));
	if(controls.calA) {
		IDispEventImpl<IDC_CALA, CMainDlg, &__uuidof(DTCtlsLibA::_ICalendarEvents), &LIBID_DTCtlsLibA, 1, 5>::DispEventAdvise(controls.calA);
	}

	pickerAWnd.SubclassWindow(GetDlgItem(IDC_PICKERA));
	pickerAWnd.QueryControl(__uuidof(DTCtlsLibA::IDateTimePicker), reinterpret_cast<LPVOID*>(&controls.pickerA));
	if(controls.pickerA) {
		IDispEventImpl<IDC_PICKERA, CMainDlg, &__uuidof(DTCtlsLibA::_IDateTimePickerEvents), &LIBID_DTCtlsLibA, 1, 5>::DispEventAdvise(controls.pickerA);
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.logEdit.IsWindow() && aboutButton.IsWindow()) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			CRect rc;
			pickerUWnd.GetWindowRect(&rc);
			ScreenToClient(&rc);
			pickerUWnd.MoveWindow(8, 8, rc.Width(), rc.Height());
			pickerUWnd.GetWindowRect(&rc);
			ScreenToClient(&rc);

			OLE_XSIZE_PIXELS minWidth = 0;
			OLE_YSIZE_PIXELS minHeight = 0;
			controls.calU->GetMinimumSize(&minWidth, &minHeight);
			calUWnd.MoveWindow(rc.right + 11, rc.top, minWidth, minHeight);
			pickerAWnd.MoveWindow(rc.left, rc.top + minHeight + 11, rc.Width(), rc.Height());
			calAWnd.MoveWindow(rc.right + 11, rc.top + minHeight + 11, minWidth, minHeight);

			CRect clientRectangle;
			GetClientRect(&clientRectangle);
			int x = rc.right + 11 + minWidth + 11;
			controls.logEdit.MoveWindow(x, 0, clientRectangle.Width() - x, clientRectangle.Height() - 32);
			aboutButton.SetWindowPos(NULL, x, clientRectangle.Height() - 27, 0, 0, SWP_NOSIZE);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.calU && GetDlgItem(IDC_CALU).IsChild(hFocusedControl)) {
		controls.calU->About();
	} else if(controls.pickerU && GetDlgItem(IDC_PICKERU).IsChild(hFocusedControl)) {
		controls.pickerU->About();
	} else if(controls.calA && GetDlgItem(IDC_CALA).IsChild(hFocusedControl)) {
		controls.calA->About();
	} else if(controls.pickerA && GetDlgItem(IDC_PICKERA).IsChild(hFocusedControl)) {
		controls.pickerA->About();
	}
	return 0;
}

LRESULT CMainDlg::OnSetFocusAbout(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled)
{
	hFocusedControl = reinterpret_cast<HWND>(wParam);
	bHandled = FALSE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void __stdcall CMainDlg::ClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_Click: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_ContextMenu: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), date, button, shift, x, y, hitTestDetails, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DateMouseEnterCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_DateMouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DateMouseLeaveCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_DateMouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_DblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowCalU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CalU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetBoldDatesCalU(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(firstDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_GetBoldDates: firstDate=%s, numberOfDates=%i"), date, numberOfDates);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownCalU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CalU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressCalU(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("CalU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpCalU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CalU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MDblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MouseDown: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MouseHover: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MouseMove: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MouseUp: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_MouseWheel: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i, hitTestDetails=0x%X"), date, button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropCalU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), *effect, date, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.calU->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterCalU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveCalU(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalU_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveCalU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_RClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_RDblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowCalU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CalU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowCalU()
{
	AddLogEntry(CAtlString(TEXT("CalU_ResizedControlWindow")));
}

void __stdcall CMainDlg::SelectionChangedCalU(DATE firstSelectedDate, DATE lastSelectedDate)
{
	CAtlString str;
	CComBSTR date1;
	VarBstrFromDate(firstSelectedDate, 0, LOCALE_NOUSEROVERRIDE, &date1);
	CComBSTR date2;
	VarBstrFromDate(lastSelectedDate, 0, LOCALE_NOUSEROVERRIDE, &date2);
	str.Format(TEXT("CalU_SelectionChanged: firstSelectedDate=%s, lastSelectedDate=%s"), date1, date2);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ViewChangedCalU(DTCtlsLibU::ViewConstants oldView, DTCtlsLibU::ViewConstants newView)
{
	CAtlString str;
	str.Format(TEXT("CalU_ViewChanged: oldView=%i, newView=%i"), oldView, newView);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_XClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalU_XDblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarCloseUpPickerU()
{
	AddLogEntry(TEXT("PickerU_CalendarCloseUp"));
}

void __stdcall CMainDlg::CalendarContextMenuPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarContextMenu: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), date, button, shift, x, y, hitTestDetails, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarDateMouseEnterPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarDateMouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarDateMouseLeavePickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarDateMouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarDropDownPickerU()
{
	AddLogEntry(TEXT("PickerU_CalendarDropDown"));
}

void __stdcall CMainDlg::CalendarKeyDownPickerU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerU_CalendarKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarKeyPressPickerU(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("PickerU_CalendarKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarKeyUpPickerU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerU_CalendarKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseDownPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMouseDown: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseEnterPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseHoverPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMouseHover: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseLeavePickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseMovePickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMouseMove: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseUpPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMouseUp: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseWheelPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarMouseWheel: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i, hitTestDetails=0x%X"), date, button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarOLEDragDropPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_CalendarOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_CalendarOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), *effect, date, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.pickerU->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::CalendarOLEDragEnterPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_CalendarOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_CalendarOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarOLEDragLeavePickerU(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_CalendarOLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoCloseUp=%i"), date, button, shift, x, y, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarOLEDragMouseMovePickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_CalendarOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarRClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarRClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarViewChangedPickerU(DTCtlsLibU::ViewConstants oldView, DTCtlsLibU::ViewConstants newView)
{
	CAtlString str;
	str.Format(TEXT("PickerU_CalendarViewChanged: oldView=%i, newView=%i"), oldView, newView);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarXClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CalendarXClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CallbackFieldKeyDownPickerU(BSTR callbackField, short keyCode, short shift, DATE* currentDate)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(*currentDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CallbackFieldKeyDown: callbackField=%s, keyCode=%i, shift=%i, currentDate=%s"), callbackField, keyCode, shift, date);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuPickerU(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("PickerU_ContextMenu: button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedCalendarControlWindowPickerU(long hWndCalendar)
{
	CAtlString str;
	str.Format(TEXT("PickerU_CreatedCalendarControlWindow: hWndCalendar=0x%X"), hWndCalendar);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowPickerU(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("PickerU_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CurrentDateChangedPickerU(VARIANT_BOOL dateSelected, DATE selectedDate)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(selectedDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_CurrentDateChanged: dateSelected=%i, selectedDate=%s"), dateSelected, date);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedCalendarControlWindowPickerU(long hWndCalendar)
{
	CAtlString str;
	str.Format(TEXT("PickerU_DestroyedCalendarControlWindow: hWndCalendar=0x%X"), hWndCalendar);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowPickerU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("PickerU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowPickerU(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("PickerU_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetBoldDatesPickerU(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(firstDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_GetBoldDates: firstDate=%s, numberOfDates=%i"), date, numberOfDates);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetCallbackFieldTextPickerU(BSTR callbackField, DATE dateToFormat, BSTR* textToDisplay)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(dateToFormat, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_GetCallbackFieldText: callbackField=%s, dateToFormat=%s, textToDisplay=%s"), callbackField, date, *textToDisplay);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetCallbackFieldTextSizePickerU(BSTR callbackField, OLE_XSIZE_PIXELS* textWidth, OLE_YSIZE_PIXELS* textHeight)
{
	CAtlString str;
	str.Format(TEXT("PickerU_GetCallbackFieldTextSize: callbackField=%s, textWidth=%i, textHeight=%i"), callbackField, *textWidth, *textHeight);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownPickerU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressPickerU(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("PickerU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpPickerU(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeavePickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMovePickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelPickerU(short button, short shift, long x, long y, DTCtlsLibU::ScrollAxisConstants scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("PickerU_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.pickerU->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoDropDown=%i"), *effect, button, shift, x, y, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePickerU(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMovePickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoDropDown=%i"), *effect, button, shift, x, y, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ParseUserInputPickerU(BSTR userInput, DATE* inputDate, VARIANT_BOOL* dateIsValid)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(*inputDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerU_ParseUserInput: userInput=%s, inputDate=%s, dateIsValid=%i"), userInput, date, *dateIsValid);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowPickerU(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("PickerU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowPickerU()
{
	AddLogEntry(CAtlString(TEXT("PickerU_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickPickerU(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerU_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_Click: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_ContextMenu: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), date, button, shift, x, y, hitTestDetails, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DateMouseEnterCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_DateMouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DateMouseLeaveCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_DateMouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_DblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowCalA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CalA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetBoldDatesCalA(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(firstDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_GetBoldDates: firstDate=%s, numberOfDates=%i"), date, numberOfDates);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownCalA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CalA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressCalA(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("CalA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpCalA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("CalA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MDblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MouseDown: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MouseHover: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MouseMove: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MouseUp: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_MouseWheel: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i, hitTestDetails=0x%X"), date, button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropCalA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), *effect, date, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.calA->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterCalA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveCalA(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalA_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveCalA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("CalA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("CalA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_RClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_RDblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowCalA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("CalA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowCalA()
{
	AddLogEntry(CAtlString(TEXT("CalA_ResizedControlWindow")));
}

void __stdcall CMainDlg::SelectionChangedCalA(DATE firstSelectedDate, DATE lastSelectedDate)
{
	CAtlString str;
	CComBSTR date1;
	VarBstrFromDate(firstSelectedDate, 0, LOCALE_NOUSEROVERRIDE, &date1);
	CComBSTR date2;
	VarBstrFromDate(lastSelectedDate, 0, LOCALE_NOUSEROVERRIDE, &date2);
	str.Format(TEXT("CalA_SelectionChanged: firstSelectedDate=%s, lastSelectedDate=%s"), date1, date2);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ViewChangedCalA(DTCtlsLibA::ViewConstants oldView, DTCtlsLibA::ViewConstants newView)
{
	CAtlString str;
	str.Format(TEXT("CalA_ViewChanged: oldView=%i, newView=%i"), oldView, newView);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_XClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("CalA_XDblClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarCloseUpPickerA()
{
	AddLogEntry(TEXT("PickerA_CalendarCloseUp"));
}

void __stdcall CMainDlg::CalendarContextMenuPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarContextMenu: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, showDefaultMenu=%i"), date, button, shift, x, y, hitTestDetails, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarDateMouseEnterPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarDateMouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarDateMouseLeavePickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarDateMouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarDropDownPickerA()
{
	AddLogEntry(TEXT("PickerA_CalendarDropDown"));
}

void __stdcall CMainDlg::CalendarKeyDownPickerA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerA_CalendarKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarKeyPressPickerA(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("PickerA_CalendarKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarKeyUpPickerA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerA_CalendarKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseDownPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMouseDown: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseEnterPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMouseEnter: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseHoverPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMouseHover: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseLeavePickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMouseLeave: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseMovePickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMouseMove: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseUpPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMouseUp: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarMouseWheelPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarMouseWheel: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i, hitTestDetails=0x%X"), date, button, shift, x, y, scrollAxis, wheelDelta, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarOLEDragDropPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_CalendarOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_CalendarOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), *effect, date, button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.pickerA->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::CalendarOLEDragEnterPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_CalendarOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_CalendarOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarOLEDragLeavePickerA(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_CalendarOLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoCloseUp=%i"), date, button, shift, x, y, hitTestDetails, *autoCloseUp);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarOLEDragMouseMovePickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_CalendarOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	CComBSTR date;
	VarBstrFromDate(dropTarget, 0, LOCALE_NOUSEROVERRIDE, &date);
	tmp.Format(TEXT(", effect=%i, dropTarget=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X, autoHScrollVelocity=%i"), *effect, date, button, shift, x, y, hitTestDetails, *autoHScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarRClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarRClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarViewChangedPickerA(DTCtlsLibA::ViewConstants oldView, DTCtlsLibA::ViewConstants newView)
{
	CAtlString str;
	str.Format(TEXT("PickerA_CalendarViewChanged: oldView=%i, newView=%i"), oldView, newView);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CalendarXClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(hitDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CalendarXClick: hitDate=%s, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=0x%X"), date, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CallbackFieldKeyDownPickerA(BSTR callbackField, short keyCode, short shift, DATE* currentDate)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(*currentDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CallbackFieldKeyDown: callbackField=%s, keyCode=%i, shift=%i, currentDate=%s"), callbackField, keyCode, shift, date);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuPickerA(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("PickerA_ContextMenu: button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedCalendarControlWindowPickerA(long hWndCalendar)
{
	CAtlString str;
	str.Format(TEXT("PickerA_CreatedCalendarControlWindow: hWndCalendar=0x%X"), hWndCalendar);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowPickerA(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("PickerA_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CurrentDateChangedPickerA(VARIANT_BOOL dateSelected, DATE selectedDate)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(selectedDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_CurrentDateChanged: dateSelected=%i, selectedDate=%s"), dateSelected, date);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedCalendarControlWindowPickerA(long hWndCalendar)
{
	CAtlString str;
	str.Format(TEXT("PickerA_DestroyedCalendarControlWindow: hWndCalendar=0x%X"), hWndCalendar);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowPickerA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("PickerA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowPickerA(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("PickerA_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetBoldDatesPickerA(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(firstDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_GetBoldDates: firstDate=%s, numberOfDates=%i"), date, numberOfDates);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetCallbackFieldTextPickerA(BSTR callbackField, DATE dateToFormat, BSTR* textToDisplay)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(dateToFormat, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_GetCallbackFieldText: callbackField=%s, dateToFormat=%s, textToDisplay=%s"), callbackField, date, *textToDisplay);

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetCallbackFieldTextSizePickerA(BSTR callbackField, OLE_XSIZE_PIXELS* textWidth, OLE_YSIZE_PIXELS* textHeight)
{
	CAtlString str;
	str.Format(TEXT("PickerA_GetCallbackFieldTextSize: callbackField=%s, textWidth=%i, textHeight=%i"), callbackField, *textWidth, *textHeight);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownPickerA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyPressPickerA(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("PickerA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpPickerA(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("PickerA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeavePickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMovePickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseWheelPickerA(short button, short shift, long x, long y, DTCtlsLibA::ScrollAxisConstants scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("PickerA_MouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=0x%X, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData) {
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				str = TEXT("");
				for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
					str += array.GetAt(i);
					str += TEXT("\r\n");
				}
				controls.pickerA->FinishOLEDragDrop();
				MessageBox(str, TEXT("Dropped files"));
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoDropDown=%i"), *effect, button, shift, x, y, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePickerA(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_OLEDragLeave: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMovePickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown)
{
	CAtlString str;
	CComQIPtr<DTCtlsLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("PickerA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("PickerA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i, autoDropDown=%i"), *effect, button, shift, x, y, *autoDropDown);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ParseUserInputPickerA(BSTR userInput, DATE* inputDate, VARIANT_BOOL* dateIsValid)
{
	CAtlString str;
	CComBSTR date;
	VarBstrFromDate(*inputDate, 0, LOCALE_NOUSEROVERRIDE, &date);
	str.Format(TEXT("PickerA_ParseUserInput: userInput=%s, inputDate=%s, dateIsValid=%i"), userInput, date, *dateIsValid);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowPickerA(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("PickerA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowPickerA()
{
	AddLogEntry(CAtlString(TEXT("PickerA_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickPickerA(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("PickerA_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}
