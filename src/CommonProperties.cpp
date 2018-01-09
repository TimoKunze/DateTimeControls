// CommonProperties.cpp: The controls' "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));
	controls.minDatePicker = GetDlgItem(IDC_MINDATEPICKER);
	controls.maxDatePicker = GetDlgItem(IDC_MAXDATEPICKER);

	// setup the toolbar
	WTL::CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(pNotificationDetails->hwndFrom != controls.disabledEventsList) {
		return 0;
	}

	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	int itemIndex = pDetails->iItem;
	if(properties.numberOfDateTimePickers == 0) {
		switch(itemIndex) {
			case 1:
				itemIndex = 2;
				break;
			case 2:
				itemIndex = 4;
				break;
		}
	}

	switch(itemIndex) {
		case 0:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, DateMouseEnter, DateMouseLeave, MouseMove, MouseWheel"))));
			break;
		case 1:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("CalendarMouseDown, CalendarMouseUp, CalendarMouseEnter, CalendarMouseHover, CalendarMouseLeave, CalendarMouseLeave, CalendarDateMouseEnter, CalendarDateMouseLeave, CalendarMouseMove, CalendarMouseWheel"))));
			break;
		case 2:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Click, DblClick, MClick, MDblClick, RClick, RDblClick, XClick, XDblClick"))));
			break;
		case 3:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("CalendarClick, CalendarMClick, CalendarRClick, CalendarXClick"))));
			break;
		case 4:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("KeyDown, KeyUp, KeyPress"))));
			break;
		case 5:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("CalendarKeyDown, CalendarKeyUp, CalendarKeyPress"))));
			break;
		case 6:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ParseUserInput"))));
			break;
	}
	ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT CommonProperties::OnPickerDateTimeChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	SetDirty(TRUE);
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_ICalendar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ICalendar*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IDateTimePicker, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IDateTimePicker*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_ICalendar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("Calendar Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<ICalendar*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IDateTimePicker, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("DateTimePicker Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IDateTimePicker*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_ICalendar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<ICalendar*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState((properties.numberOfDateTimePickers > 0 ? 2 : 1), LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState((properties.numberOfDateTimePickers > 0 ? 4 : 2), LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			reinterpret_cast<ICalendar*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			if(controls.minDatePicker.IsWindowEnabled()) {
				SYSTEMTIME sysTime;
				controls.minDatePicker.GetSystemTime(&sysTime);
				DATE date;
				SystemTimeToVariantTime(&sysTime, &date);
				reinterpret_cast<ICalendar*>(pControl)->put_MinDate(date);
				controls.maxDatePicker.GetSystemTime(&sysTime);
				SystemTimeToVariantTime(&sysTime, &date);
				reinterpret_cast<ICalendar*>(pControl)->put_MaxDate(date);
			}
			pControl->Release();

		} else if(m_ppUnk[object]->QueryInterface(IID_IDateTimePicker, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IDateTimePicker*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deCalendarMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deCalendarClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, deCalendarKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, deParseUserInput);
			reinterpret_cast<IDateTimePicker*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			if(controls.minDatePicker.IsWindowEnabled()) {
				SYSTEMTIME sysTime;
				controls.minDatePicker.GetSystemTime(&sysTime);
				DATE date;
				SystemTimeToVariantTime(&sysTime, &date);
				reinterpret_cast<IDateTimePicker*>(pControl)->put_MinDate(date);
				controls.maxDatePicker.GetSystemTime(&sysTime);
				SystemTimeToVariantTime(&sysTime, &date);
				reinterpret_cast<IDateTimePicker*>(pControl)->put_MaxDate(date);
			}
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));
	controls.minDatePicker.EnableWindow(m_nObjects == 1);
	controls.maxDatePicker.EnableWindow(m_nObjects == 1);

	// get the settings
	properties.numberOfDateTimePickers = 0;
	int numberOfCalendars = 0;
	DisabledEventsConstants* pDisabledEvents = static_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		ZeroMemory(pDisabledEvents, m_nObjects * sizeof(DisabledEventsConstants));
		for(UINT object = 0; object < m_nObjects; ++object) {
			LPUNKNOWN pControl = NULL;
			if(m_ppUnk[object]->QueryInterface(IID_ICalendar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++numberOfCalendars;
				reinterpret_cast<ICalendar*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();

			} else if(m_ppUnk[object]->QueryInterface(IID_IDateTimePicker, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++properties.numberOfDateTimePickers;
				reinterpret_cast<IDateTimePicker*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();
			}
		}

		// fill the listboxes
		LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
		properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
		controls.disabledEventsList.DeleteAllItems();
		controls.disabledEventsList.AddItem(0, 0, TEXT("Mouse events"));
		controls.disabledEventsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, deMouseEvents), LVIS_STATEIMAGEMASK);
		controls.disabledEventsList.AddItem(1, 0, TEXT("Click events"));
		controls.disabledEventsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, deClickEvents), LVIS_STATEIMAGEMASK);
		controls.disabledEventsList.AddItem(2, 0, TEXT("Keyboard events"));
		controls.disabledEventsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, deKeyboardEvents), LVIS_STATEIMAGEMASK);
		if(properties.numberOfDateTimePickers > 0) {
			controls.disabledEventsList.AddItem(1, 0, TEXT("Mouse events (drop-down calendar)"));
			controls.disabledEventsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, deCalendarMouseEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(3, 0, TEXT("Click events (drop-down calendar)"));
			controls.disabledEventsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, deCalendarClickEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(5, 0, TEXT("Keyboard events (drop-down calendar)"));
			controls.disabledEventsList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, deCalendarKeyboardEvents), LVIS_STATEIMAGEMASK);
			controls.disabledEventsList.AddItem(6, 0, TEXT("ParseUserInput event"));
			controls.disabledEventsList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, deParseUserInput), LVIS_STATEIMAGEMASK);
		}
		controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

		properties.hWndCheckMarksAreSetFor = NULL;

		if(m_nObjects == 1) {
			DATE minDate = 0;
			DATE maxDate = 0;

			LPUNKNOWN pControl = NULL;
			if(m_ppUnk[0]->QueryInterface(IID_ICalendar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				reinterpret_cast<ICalendar*>(pControl)->get_MinDate(&minDate);
				reinterpret_cast<ICalendar*>(pControl)->get_MaxDate(&maxDate);
				pControl->Release();

			} else if(m_ppUnk[0]->QueryInterface(IID_IDateTimePicker, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				reinterpret_cast<IDateTimePicker*>(pControl)->get_MinDate(&minDate);
				reinterpret_cast<IDateTimePicker*>(pControl)->get_MaxDate(&maxDate);
				pControl->Release();
			}

			SYSTEMTIME sysTime;
			VariantTimeToSystemTime(minDate, &sysTime);
			controls.minDatePicker.SetSystemTime(GDT_VALID, &sysTime);
			VariantTimeToSystemTime(maxDate, &sysTime);
			controls.maxDatePicker.SetSystemTime(GDT_VALID, &sysTime);
		}

		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}