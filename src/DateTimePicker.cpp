// DateTimePicker.cpp: Superclasses SysDateTimePick32.

#include "stdafx.h"
#include "DateTimePicker.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")

#define MCS_MASK (MCS_DAYSTATE | MCS_MULTISELECT | MCS_NOSELCHANGEONNAV | MCS_NOTODAY | MCS_NOTODAYCIRCLE | MCS_NOTRAILINGDATES | MCS_SHORTDAYSOFWEEK | MCS_WEEKNUMBERS)


// initialize complex constants
IMEModeConstants DateTimePicker::IMEFlags::chineseIMETable[10] = {
    imeOff,
    imeOff,
    imeOff,
    imeOff,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOff
};

IMEModeConstants DateTimePicker::IMEFlags::japaneseIMETable[10] = {
    imeDisable,
    imeDisable,
    imeOff,
    imeOff,
    imeHiragana,
    imeHiragana,
    imeKatakana,
    imeKatakanaHalf,
    imeAlphaFull,
    imeAlpha
};

IMEModeConstants DateTimePicker::IMEFlags::koreanIMETable[10] = {
    imeDisable,
    imeDisable,
    imeAlpha,
    imeAlpha,
    imeHangulFull,
    imeHangul,
    imeHangulFull,
    imeHangul,
    imeAlphaFull,
    imeAlpha
};

FONTDESC DateTimePicker::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


#pragma warning(push)
#pragma warning(disable: 4355)     // passing "this" to a member constructor
DateTimePicker::DateTimePicker() :
    containedCalendar(MONTHCAL_CLASS, this, 1),
    containedEdit(WC_EDIT, this, 2)
{
	properties.calendarFont.InitializePropertyWatcher(this, DISPID_DTP_CALENDARFONT);
	properties.font.InitializePropertyWatcher(this, DISPID_DTP_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_DTP_MOUSEICON);

	SIZEL size = {100, 20};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	dateUnderMouse = 0;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}
}
#pragma warning(pop)


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP DateTimePicker::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IDateTimePicker, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP DateTimePicker::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<DateTimePicker>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);

		CComBSTR bstr;
		hr = pPropertyBag->Read(OLESTR("CustomDateFormat"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_CustomDateFormat(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP DateTimePicker::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<DateTimePicker>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.customDateFormat.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("CustomDateFormat"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP DateTimePicker::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 21 VT_I4 properties...
	pSize->QuadPart += 21 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 16 VT_BOOL properties...
	pSize->QuadPart += 16 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 2 VT_DATE properties...
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(DATE));
	// ...and 1 VT_BSTR property...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.customDateFormat.ByteLength() + sizeof(OLECHAR);

	// ...and 3 VT_DISPATCH properties
	pSize->QuadPart += 3 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.calendarFont.pFontDisp) {
		if(FAILED(properties.calendarFont.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.calendarFont.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP DateTimePicker::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	// NOTE: We neither support legacy streams nor streams with a version < 0x0102.

	HRESULT hr = S_OK;
	LONG signature = 0;
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x0E0E0E0E/*4x AppID*/) {
		// might be a legacy stream, that starts with the ATL version
		if(signature == 0x0900) {
			version = 0x0099;
		} else {
			return E_FAIL;
		}
	}
	//LONG version = 0;
	if(version != 0x0099) {
		if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
			return hr;
		}
		if(version > 0x0102) {
			return E_FAIL;
		}
		LONG subSignature = 0;
		if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
			return hr;
		}
		if(subSignature != 0x02020202/*4x 0x02 (-> DateTimePicker)*/) {
			return E_FAIL;
		}
	}

	DWORD atlVersion;
	if(version == 0x0099) {
		atlVersion = 0x0900;
	} else {
		if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
			return hr;
		}
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version != 0x0100) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowNullSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR calendarBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG calendarDragScrollTimeBase = propertyValue.lVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pCalendarFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pCalendarFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR calendarForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL calendarHighlightTodayDate = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG calendarHoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL calendarKeepSelectionOnNavigation = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR calendarMonthBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL calendarShowTodayDate = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL calendarShowTrailingDates = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL calendarShowWeekNumbers = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR calendarTitleBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR calendarTitleForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR calendarTrailingForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL calendarUseShortestDayNames = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL calendarUseSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ViewConstants calendarView = static_cast<ViewConstants>(propertyValue.lVal);
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR customDateFormat;
	if(FAILED(hr = customDateFormat.ReadFromStream(pStream))) {
		return hr;
	}
	if(!customDateFormat) {
		customDateFormat = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DateFormatConstants dateFormat = static_cast<DateFormatConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dateSelected = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragDropDownTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DropDownAlignmentConstants dropDownAlignment = static_cast<DropDownAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants imeMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_DATE, &propertyValue))) {
		return hr;
	}
	DATE maxDate = propertyValue.date;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_DATE, &propertyValue))) {
		return hr;
	}
	DATE minDate = propertyValue.date;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	StartOfWeekConstants startOfWeek = static_cast<StartOfWeekConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	StyleConstants style = static_cast<StyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;

	VARIANT_BOOL detectDoubleClicks = VARIANT_TRUE;
	if(version >= 0x0102) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		detectDoubleClicks = propertyValue.boolVal;
	}


	hr = put_AllowNullSelection(allowNullSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarBackColor(calendarBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarDragScrollTimeBase(calendarDragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarFont(pCalendarFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarForeColor(calendarForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarHighlightTodayDate(calendarHighlightTodayDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarHoverTime(calendarHoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarKeepSelectionOnNavigation(calendarKeepSelectionOnNavigation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarMonthBackColor(calendarMonthBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarShowTodayDate(calendarShowTodayDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarShowTrailingDates(calendarShowTrailingDates);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarShowWeekNumbers(calendarShowWeekNumbers);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarTitleBackColor(calendarTitleBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarTitleForeColor(calendarTitleForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarTrailingForeColor(calendarTrailingForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarUseShortestDayNames(calendarUseShortestDayNames);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarUseSystemFont(calendarUseSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CalendarView(calendarView);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomDateFormat(customDateFormat);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DateFormat(dateFormat);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DateSelected(dateSelected);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DetectDoubleClicks(detectDoubleClicks);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragDropDownTime(dragDropDownTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DropDownAlignment(dropDownAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEMode(imeMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxDate(maxDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinDate(minDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_StartOfWeek(startOfWeek);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Style(style);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP DateTimePicker::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x0E0E0E0E/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0102;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x02020202/*4x 0x02 (-> DateTimePicker)*/;
	if(FAILED(hr = pStream->Write(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowNullSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.calendarBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.calendarDragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.calendarFont.pFontDisp) {
		if(FAILED(hr = properties.calendarFont.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.calendarForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.calendarHighlightTodayDate);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.calendarHoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.calendarKeepSelectionOnNavigation);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.calendarMonthBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.calendarShowTodayDate);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.calendarShowTrailingDates);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.calendarShowWeekNumbers);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.calendarTitleBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.calendarTitleForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.calendarTrailingForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.calendarUseShortestDayNames);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.calendarUseSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.calendarView;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.customDateFormat.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.lVal = properties.dateFormat;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dateSelected);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragDropDownTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.dropDownAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_DATE;
	propertyValue.date = properties.maxDate;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.date = properties.minDate;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.startOfWeek;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.style;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0102 starts here
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND DateTimePicker::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_DATE_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<DateTimePicker>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT DateTimePicker::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("DateTimePicker ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void DateTimePicker::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP DateTimePicker::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<DateTimePicker>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}
	if(!properties.calendarFont.pFontDisp) {
		FONTDESC defaultFont = properties.calendarFont.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_CalendarFont(pFont);
	}

	return hr;
}

STDMETHODIMP DateTimePicker::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL DateTimePicker::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
		}
	}
	return CComControl<DateTimePicker>::PreTranslateAccelerator(pMessage, hReturnValue);
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP DateTimePicker::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	HWND hWnd = WindowFromPoint(buffer);
	if(hWnd == containedCalendar || hWnd == containedEdit) {
		Raise_CalendarOLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
		if(dragDropStatus.pDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragEnter(containedCalendar, pDataObject, &buffer, *pEffect);
		}
	} else {
		Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
		if(dragDropStatus.pDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::DragLeave(void)
{
	if(dragDropStatus.isOverCalendar) {
		Raise_CalendarOLEDragLeave(FALSE);
	} else {
		Raise_OLEDragLeave(FALSE);
	}
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	HWND hWnd = WindowFromPoint(buffer);
	if(hWnd == containedCalendar || hWnd == containedEdit) {
		if(dragDropStatus.isOverCalendar) {
			Raise_CalendarOLEDragMouseMove(pEffect, keyState, mousePosition);
		} else {
			// we've left the picker and entered the calendar
			Raise_OLEDragLeave(TRUE);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			IDataObject_GetDropDescription(pDataObject, newDropDescription);
			if(memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_CalendarOLEDragEnter(NULL, pEffect, keyState, mousePosition);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(containedCalendar, pDataObject, &buffer, *pEffect);
			}
		}
	} else {
		if(dragDropStatus.isOverCalendar) {
			// we've left the calendar and entered the picker
			Raise_CalendarOLEDragLeave(TRUE);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			IDataObject_GetDropDescription(pDataObject, newDropDescription);
			if(memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_OLEDragEnter(NULL, pEffect, keyState, mousePosition);
			if(dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			}
		} else {
			Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
		}
	}
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	HWND hWnd = WindowFromPoint(buffer);
	if(hWnd == containedCalendar || hWnd == containedEdit) {
		Raise_CalendarOLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	} else {
		Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	}
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP DateTimePicker::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Accessibility:
			*pName = GetResString(IDPC_ACCESSIBILITY).Detach();
			return S_OK;
			break;
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_DTP_CHECKBOXOBJECTSTATE:
		case DISPID_DTP_DROPDOWNBUTTONOBJECTSTATE:
			*pCategory = PROPCAT_Accessibility;
			return S_OK;
			break;
		case DISPID_DTP_ALLOWNULLSELECTION:
		case DISPID_DTP_APPEARANCE:
		case DISPID_DTP_BORDERSTYLE:
		case DISPID_DTP_CALENDARGRIDCELLSELECTED:
		case DISPID_DTP_CALENDARHIGHLIGHTTODAYDATE:
		case DISPID_DTP_CALENDARSHOWTODAYDATE:
		case DISPID_DTP_CALENDARSHOWTRAILINGDATES:
		case DISPID_DTP_CALENDARSHOWWEEKNUMBERS:
		case DISPID_DTP_CALENDARUSESHORTESTDAYNAMES:
		case DISPID_DTP_CALENDARVIEW:
		case DISPID_DTP_CUSTOMDATEFORMAT:
		case DISPID_DTP_DATEFORMAT:
		case DISPID_DTP_DROPDOWNALIGNMENT:
		case DISPID_DTP_MOUSEICON:
		case DISPID_DTP_MOUSEPOINTER:
		case DISPID_DTP_STARTOFWEEK:
		case DISPID_DTP_STYLE:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_DTP_CALENDARHOVERTIME:
		case DISPID_DTP_CALENDARKEEPSELECTIONONNAVIGATION:
		case DISPID_DTP_DETECTDOUBLECLICKS:
		case DISPID_DTP_DISABLEDEVENTS:
		case DISPID_DTP_DONTREDRAW:
		case DISPID_DTP_HOVERTIME:
		case DISPID_DTP_IMEMODE:
		case DISPID_DTP_MAXDATE:
		case DISPID_DTP_MINDATE:
		case DISPID_DTP_PROCESSCONTEXTMENUKEYS:
		case DISPID_DTP_RIGHTTOLEFT:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_DTP_CALENDARBACKCOLOR:
		case DISPID_DTP_CALENDARFORECOLOR:
		case DISPID_DTP_CALENDARMONTHBACKCOLOR:
		case DISPID_DTP_CALENDARTITLEBACKCOLOR:
		case DISPID_DTP_CALENDARTITLEFORECOLOR:
		case DISPID_DTP_CALENDARTRAILINGFORECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_DTP_APPID:
		case DISPID_DTP_APPNAME:
		case DISPID_DTP_APPSHORTNAME:
		case DISPID_DTP_BUILD:
		case DISPID_DTP_CHARSET:
		case DISPID_DTP_HWND:
		case DISPID_DTP_HWNDCALENDAR:
		case DISPID_DTP_HWNDEDIT:
		case DISPID_DTP_HWNDUPDOWN:
		case DISPID_DTP_ISRELEASE:
		case DISPID_DTP_PROGRAMMER:
		case DISPID_DTP_TESTER:
		case DISPID_DTP_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_DTP_CALENDARDRAGSCROLLTIMEBASE:
		case DISPID_DTP_DRAGDROPDOWNTIME:
		case DISPID_DTP_REGISTERFOROLEDRAGDROP:
		case DISPID_DTP_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_DTP_CALENDARFONT:
		case DISPID_DTP_CALENDARUSESYSTEMFONT:
		case DISPID_DTP_FONT:
		case DISPID_DTP_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_DTP_CALENDARCARETDATETEXT:
		case DISPID_DTP_CALENDARFIRSTDATE:
		case DISPID_DTP_CALENDARGRIDCELLTEXT:
		case DISPID_DTP_CALENDARHEADERTEXT:
		case DISPID_DTP_CALENDARLASTDATE:
		case DISPID_DTP_CALENDARTODAYDATE:
		case DISPID_DTP_CURRENTDATE:
		case DISPID_DTP_DATESELECTED:
		case DISPID_DTP_ENABLED:
		case DISPID_DTP_VALUE:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString DateTimePicker::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString DateTimePicker::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString DateTimePicker::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=DateTimeControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString DateTimePicker::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString DateTimePicker::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString DateTimePicker::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP DateTimePicker::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_DTP_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_DTP_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_DTP_CALENDARVIEW:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.calendarView), description);
			break;
		case DISPID_DTP_CUSTOMDATEFORMAT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_DTP_DATEFORMAT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.dateFormat), description);
			break;
		case DISPID_DTP_DROPDOWNALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.dropDownAlignment), description);
			break;
		case DISPID_DTP_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_DTP_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_DTP_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_DTP_STARTOFWEEK:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.startOfWeek), description);
			break;
		case DISPID_DTP_STYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.style), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<DateTimePicker>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP DateTimePicker::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_DTP_BORDERSTYLE:
		case DISPID_DTP_DROPDOWNALIGNMENT:
		case DISPID_DTP_STYLE:
			c = 2;
			break;
		case DISPID_DTP_APPEARANCE:
		case DISPID_DTP_CALENDARVIEW:
		case DISPID_DTP_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_DTP_DATEFORMAT:
			c = 5;
			break;
		case DISPID_DTP_STARTOFWEEK:
			c = 8;
			break;
		case DISPID_DTP_IMEMODE:
			c = 12;
			break;
		case DISPID_DTP_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<DateTimePicker>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_DTP_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if(property == DISPID_DTP_IMEMODE || property == DISPID_DTP_STARTOFWEEK) {
			// the enum is -1-based
			--propertyValue;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_DTP_APPEARANCE:
		case DISPID_DTP_BORDERSTYLE:
		case DISPID_DTP_CALENDARVIEW:
		case DISPID_DTP_DATEFORMAT:
		case DISPID_DTP_DROPDOWNALIGNMENT:
		case DISPID_DTP_IMEMODE:
		case DISPID_DTP_MOUSEPOINTER:
		case DISPID_DTP_RIGHTTOLEFT:
		case DISPID_DTP_STARTOFWEEK:
		case DISPID_DTP_STYLE:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<DateTimePicker>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_DTP_CUSTOMDATEFORMAT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<DateTimePicker>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT DateTimePicker::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_DTP_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				case aDefault:
					description = GetResStringWithNumber(IDP_APPEARANCEDEFAULT, aDefault);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_CALENDARVIEW:
			switch(cookie) {
				case vMonth:
					description = GetResStringWithNumber(IDP_VIEWMONTH, vMonth);
					break;
				case vYear:
					description = GetResStringWithNumber(IDP_VIEWYEAR, vYear);
					break;
				case vDecade:
					description = GetResStringWithNumber(IDP_VIEWDECADE, vDecade);
					break;
				case vCentury:
					description = GetResStringWithNumber(IDP_VIEWCENTURY, vCentury);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_DATEFORMAT:
			switch(cookie) {
				case dfShortDate:
					description = GetResStringWithNumber(IDP_DATEFORMATSHORTDATE, dfShortDate);
					break;
				case dfShortDateWithFourDigitYear:
					description = GetResStringWithNumber(IDP_DATEFORMATSHORTDATE4DIGITS, dfShortDateWithFourDigitYear);
					break;
				case dfLongDate:
					description = GetResStringWithNumber(IDP_DATEFORMATLONGDATE, dfLongDate);
					break;
				case dfTime:
					description = GetResStringWithNumber(IDP_DATEFORMATTIME, dfTime);
					break;
				case dfCustom:
					description = GetResStringWithNumber(IDP_DATEFORMATCUSTOM, dfCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_DROPDOWNALIGNMENT:
			switch(cookie) {
				case ddaLeft:
					description = GetResStringWithNumber(IDP_DROPDOWNALIGNMENTLEFT, ddaLeft);
					break;
				case ddaRight:
					description = GetResStringWithNumber(IDP_DROPDOWNALIGNMENTRIGHT, ddaRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_IMEMODE:
			switch(cookie) {
				case imeInherit:
					description = GetResStringWithNumber(IDP_IMEMODEINHERIT, imeInherit);
					break;
				case imeNoControl:
					description = GetResStringWithNumber(IDP_IMEMODENOCONTROL, imeNoControl);
					break;
				case imeOn:
					description = GetResStringWithNumber(IDP_IMEMODEON, imeOn);
					break;
				case imeOff:
					description = GetResStringWithNumber(IDP_IMEMODEOFF, imeOff);
					break;
				case imeDisable:
					description = GetResStringWithNumber(IDP_IMEMODEDISABLE, imeDisable);
					break;
				case imeHiragana:
					description = GetResStringWithNumber(IDP_IMEMODEHIRAGANA, imeHiragana);
					break;
				case imeKatakana:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANA, imeKatakana);
					break;
				case imeKatakanaHalf:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANAHALF, imeKatakanaHalf);
					break;
				case imeAlphaFull:
					description = GetResStringWithNumber(IDP_IMEMODEALPHAFULL, imeAlphaFull);
					break;
				case imeAlpha:
					description = GetResStringWithNumber(IDP_IMEMODEALPHA, imeAlpha);
					break;
				case imeHangulFull:
					description = GetResStringWithNumber(IDP_IMEMODEHANGULFULL, imeHangulFull);
					break;
				case imeHangul:
					description = GetResStringWithNumber(IDP_IMEMODEHANGUL, imeHangul);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_STARTOFWEEK:
			switch(cookie) {
				case sowUseLocale:
					description = GetResStringWithNumber(IDP_STARTOFWEEKUSELOCALE, sowUseLocale);
					break;
				case sowMonday:
					description = GetResStringWithNumber(IDP_STARTOFWEEKMONDAY, sowMonday);
					break;
				case sowTuesday:
					description = GetResStringWithNumber(IDP_STARTOFWEEKTUESDAY, sowTuesday);
					break;
				case sowWednesday:
					description = GetResStringWithNumber(IDP_STARTOFWEEKWEDNESDAY, sowWednesday);
					break;
				case sowThursday:
					description = GetResStringWithNumber(IDP_STARTOFWEEKTHURSDAY, sowThursday);
					break;
				case sowFriday:
					description = GetResStringWithNumber(IDP_STARTOFWEEKFRIDAY, sowFriday);
					break;
				case sowSaturday:
					description = GetResStringWithNumber(IDP_STARTOFWEEKSATURDAY, sowSaturday);
					break;
				case sowSunday:
					description = GetResStringWithNumber(IDP_STARTOFWEEKSUNDAY, sowSunday);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_DTP_STYLE:
			switch(cookie) {
				case sDropDown:
					description = GetResStringWithNumber(IDP_STYLEDROPDOWN, sDropDown);
					break;
				case sUpDown:
					description = GetResStringWithNumber(IDP_STYLEUPDOWN, sUpDown);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP DateTimePicker::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 5;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_StockColorPage;
		pPropertyPages->pElems[3] = CLSID_StockFontPage;
		pPropertyPages->pElems[4] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP DateTimePicker::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<DateTimePicker>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP DateTimePicker::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP DateTimePicker::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT DateTimePicker::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT DateTimePicker::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_DTP_CALENDARFONT:
			if(!properties.calendarUseSystemFont) {
				ApplyCalendarFont();
			}
			break;
		case DISPID_DTP_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT DateTimePicker::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP DateTimePicker::get_AllowNullSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowNullSelection = ((GetStyle() & DTS_SHOWNONE) == DTS_SHOWNONE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowNullSelection);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_AllowNullSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_ALLOWNULLSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowNullSelection != b) {
		properties.allowNullSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			/*if(properties.allowNullSelection) {
				ModifyStyle(0, DTS_SHOWNONE);
			} else {
				ModifyStyle(DTS_SHOWNONE, 0);
			}*/
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_DTP_ALLOWNULLSELECTION);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_APPEARANCE);
	if(newValue < 0 || newValue > aDefault) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(newValue == aDefault && !IsInDesignMode() && IsWindow()) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_DTP_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 14;
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"DateTimeControls");
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"DTCtls");
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_DTP_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(DTM_GETMCCOLOR, MCSC_BACKGROUND, 0));
		if(color != OLECOLOR2COLORREF(properties.calendarBackColor)) {
			properties.calendarBackColor = color;
		}
	}

	*pValue = properties.calendarBackColor;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARBACKCOLOR);
	if(properties.calendarBackColor != newValue) {
		properties.calendarBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(DTM_SETMCCOLOR, MCSC_BACKGROUND, OLECOLOR2COLORREF(properties.calendarBackColor));
		}
		FireOnChanged(DISPID_DTP_CALENDARBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarCaretDateText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
		WCHAR pBuffer[300];

		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDAR;
		gridInfo.dwFlags = MCGIF_NAME;
		gridInfo.cchName = 300;
		gridInfo.pszName = pBuffer;
		if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = _bstr_t(gridInfo.pszName).Detach();
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_CalendarDragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.calendarDragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarDragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARDRAGSCROLLTIMEBASE);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.calendarDragScrollTimeBase != newValue) {
		properties.calendarDragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_CALENDARDRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarFirstDate(LONG rowIndex/* = -1*/, DATE* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = (rowIndex == -1 ? MCGIP_CALENDAR : MCGIP_CALENDARROW);
		gridInfo.dwFlags = MCGIF_DATE;
		gridInfo.iRow = (rowIndex == -1 ? 0 : rowIndex);
		if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			SystemTimeToVariantTime(&gridInfo.stStart, pValue);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_CalendarFont(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.calendarFont.pFontDisp) {
		properties.calendarFont.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarFont(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARFONT);
	if(properties.calendarFont.pFontDisp != pNewFont) {
		properties.calendarFont.StopWatching();
		if(properties.calendarFont.pFontDisp) {
			properties.calendarFont.pFontDisp->Release();
			properties.calendarFont.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.calendarFont.pFontDisp));
				}
			}
		}
		properties.calendarFont.StartWatching();
	}
	if(!properties.calendarUseSystemFont) {
		ApplyCalendarFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DTP_CALENDARFONT);
	return S_OK;
}

STDMETHODIMP DateTimePicker::putref_CalendarFont(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARFONT);
	if(properties.calendarFont.pFontDisp != pNewFont) {
		properties.calendarFont.StopWatching();
		if(properties.calendarFont.pFontDisp) {
			properties.calendarFont.pFontDisp->Release();
			properties.calendarFont.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.calendarFont.pFontDisp));
		}
		properties.calendarFont.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.calendarUseSystemFont) {
		ApplyCalendarFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DTP_CALENDARFONT);
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(DTM_GETMCCOLOR, MCSC_TEXT, 0));
		if(color != OLECOLOR2COLORREF(properties.calendarForeColor)) {
			properties.calendarForeColor = color;
		}
	}

	*pValue = properties.calendarForeColor;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARFORECOLOR);
	if(properties.calendarForeColor != newValue) {
		properties.calendarForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(DTM_SETMCCOLOR, MCSC_TEXT, OLECOLOR2COLORREF(properties.calendarForeColor));
		}
		FireOnChanged(DISPID_DTP_CALENDARFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarGridCellSelected(LONG columnIndex, LONG rowIndex, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDARCELL;
		gridInfo.dwFlags = MCGIF_DATE;
		gridInfo.iCol = columnIndex;
		gridInfo.iRow = rowIndex;
		if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = BOOL2VARIANTBOOL(gridInfo.bSelected);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_CalendarGridCellText(LONG columnIndex, LONG rowIndex, BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
		WCHAR pBuffer[50];

		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDARCELL;
		gridInfo.dwFlags = MCGIF_NAME;
		gridInfo.iCol = columnIndex;
		gridInfo.iRow = rowIndex;
		gridInfo.cchName = 50;
		gridInfo.pszName = pBuffer;
		if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = _bstr_t(gridInfo.pszName).Detach();
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_CalendarHeaderText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
		WCHAR pBuffer[300];

		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDARHEADER;
		gridInfo.dwFlags = MCGIF_NAME;
		gridInfo.cchName = 300;
		gridInfo.pszName = pBuffer;
		if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = _bstr_t(gridInfo.pszName).Detach();
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_CalendarHighlightTodayDate(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.calendarHighlightTodayDate = !(SendMessage(DTM_GETMCSTYLE, 0, 0) & MCS_NOTODAYCIRCLE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.calendarHighlightTodayDate);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarHighlightTodayDate(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARHIGHLIGHTTODAYDATE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.calendarHighlightTodayDate != b) {
		properties.calendarHighlightTodayDate = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD style = SendMessage(DTM_GETMCSTYLE, 0, 0);
			if(properties.calendarHighlightTodayDate) {
				style &= ~MCS_NOTODAYCIRCLE;
			} else {
				style |= MCS_NOTODAYCIRCLE;
			}
			SendMessage(DTM_SETMCSTYLE, 0, style);
		}
		FireOnChanged(DISPID_DTP_CALENDARHIGHLIGHTTODAYDATE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarHoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.calendarHoverTime;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarHoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARHOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.calendarHoverTime != newValue) {
		properties.calendarHoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_CALENDARHOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarKeepSelectionOnNavigation(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.calendarKeepSelectionOnNavigation = ((SendMessage(DTM_GETMCSTYLE, 0, 0) & MCS_NOSELCHANGEONNAV) == MCS_NOSELCHANGEONNAV);
	}

	*pValue = BOOL2VARIANTBOOL(properties.calendarKeepSelectionOnNavigation);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarKeepSelectionOnNavigation(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARKEEPSELECTIONONNAVIGATION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.calendarKeepSelectionOnNavigation != b) {
		properties.calendarKeepSelectionOnNavigation = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			DWORD style = SendMessage(DTM_GETMCSTYLE, 0, 0);
			if(properties.calendarKeepSelectionOnNavigation) {
				style |= MCS_NOSELCHANGEONNAV;
			} else {
				style &= ~MCS_NOSELCHANGEONNAV;
			}
			SendMessage(DTM_SETMCSTYLE, 0, style);
		}
		FireOnChanged(DISPID_DTP_CALENDARKEEPSELECTIONONNAVIGATION);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarLastDate(LONG rowIndex/* = -1*/, DATE* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = (rowIndex == -1 ? MCGIP_CALENDAR : MCGIP_CALENDARROW);
		gridInfo.dwFlags = MCGIF_DATE;
		gridInfo.iRow = (rowIndex == -1 ? 0 : rowIndex);
		if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			SystemTimeToVariantTime(&gridInfo.stEnd, pValue);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_CalendarMonthBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(DTM_GETMCCOLOR, MCSC_MONTHBK, 0));
		if(color != OLECOLOR2COLORREF(properties.calendarMonthBackColor)) {
			properties.calendarMonthBackColor = color;
		}
	}

	*pValue = properties.calendarMonthBackColor;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarMonthBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARMONTHBACKCOLOR);
	if(properties.calendarMonthBackColor != newValue) {
		properties.calendarMonthBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(DTM_SETMCCOLOR, MCSC_MONTHBK, OLECOLOR2COLORREF(properties.calendarMonthBackColor));
		}
		FireOnChanged(DISPID_DTP_CALENDARMONTHBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarShowTodayDate(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.calendarShowTodayDate = !(SendMessage(DTM_GETMCSTYLE, 0, 0) & MCS_NOTODAY);
	}

	*pValue = BOOL2VARIANTBOOL(properties.calendarShowTodayDate);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarShowTodayDate(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARSHOWTODAYDATE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.calendarShowTodayDate != b) {
		properties.calendarShowTodayDate = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD style = SendMessage(DTM_GETMCSTYLE, 0, 0);
			if(properties.calendarShowTodayDate) {
				style &= ~MCS_NOTODAY;
			} else {
				style |= MCS_NOTODAY;
			}
			SendMessage(DTM_SETMCSTYLE, 0, style);
		}
		FireOnChanged(DISPID_DTP_CALENDARSHOWTODAYDATE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarShowTrailingDates(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.calendarShowTrailingDates = !(SendMessage(DTM_GETMCSTYLE, 0, 0) & MCS_NOTRAILINGDATES);
	}

	*pValue = BOOL2VARIANTBOOL(properties.calendarShowTrailingDates);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarShowTrailingDates(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARSHOWTRAILINGDATES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.calendarShowTrailingDates != b) {
		properties.calendarShowTrailingDates = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			DWORD style = SendMessage(DTM_GETMCSTYLE, 0, 0);
			if(properties.calendarShowTrailingDates) {
				style &= ~MCS_NOTRAILINGDATES;
			} else {
				style |= MCS_NOTRAILINGDATES;
			}
			SendMessage(DTM_SETMCSTYLE, 0, style);
		}
		FireOnChanged(DISPID_DTP_CALENDARSHOWTRAILINGDATES);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarShowWeekNumbers(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.calendarShowWeekNumbers = ((SendMessage(DTM_GETMCSTYLE, 0, 0) & MCS_WEEKNUMBERS) == MCS_WEEKNUMBERS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.calendarShowWeekNumbers);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarShowWeekNumbers(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARSHOWWEEKNUMBERS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.calendarShowWeekNumbers != b) {
		properties.calendarShowWeekNumbers = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD style = SendMessage(DTM_GETMCSTYLE, 0, 0);
			if(properties.calendarShowWeekNumbers) {
				style |= MCS_WEEKNUMBERS;
			} else {
				style &= ~MCS_WEEKNUMBERS;
			}
			SendMessage(DTM_SETMCSTYLE, 0, style);
		}
		FireOnChanged(DISPID_DTP_CALENDARSHOWWEEKNUMBERS);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarTitleBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(DTM_GETMCCOLOR, MCSC_TITLEBK, 0));
		if(color != OLECOLOR2COLORREF(properties.calendarTitleBackColor)) {
			properties.calendarTitleBackColor = color;
		}
	}

	*pValue = properties.calendarTitleBackColor;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarTitleBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARTITLEBACKCOLOR);
	if(properties.calendarTitleBackColor != newValue) {
		properties.calendarTitleBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(DTM_SETMCCOLOR, MCSC_TITLEBK, OLECOLOR2COLORREF(properties.calendarTitleBackColor));
		}
		FireOnChanged(DISPID_DTP_CALENDARTITLEBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarTitleForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(DTM_GETMCCOLOR, MCSC_TITLETEXT, 0));
		if(color != OLECOLOR2COLORREF(properties.calendarTitleForeColor)) {
			properties.calendarTitleForeColor = color;
		}
	}

	*pValue = properties.calendarTitleForeColor;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarTitleForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARTITLEFORECOLOR);
	if(properties.calendarTitleForeColor != newValue) {
		properties.calendarTitleForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(DTM_SETMCCOLOR, MCSC_TITLETEXT, OLECOLOR2COLORREF(properties.calendarTitleForeColor));
		}
		FireOnChanged(DISPID_DTP_CALENDARTITLEFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarTodayDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	if(containedCalendar.IsWindow()) {
		SYSTEMTIME systemTime = {0};
		if(containedCalendar.SendMessage(MCM_GETTODAY, 0, reinterpret_cast<LPARAM>(&systemTime))) {
			SystemTimeToVariantTime(&systemTime, pValue);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::put_CalendarTodayDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARTODAYDATE);
	HRESULT hr = E_FAIL;
	if(containedCalendar.IsWindow()) {
		if(newValue == 0) {
			containedCalendar.SendMessage(MCM_SETTODAY, 0, NULL);
		} else {
			SYSTEMTIME systemTime = {0};
			VariantTimeToSystemTime(newValue, &systemTime);
			containedCalendar.SendMessage(MCM_SETTODAY, 0, reinterpret_cast<LPARAM>(&systemTime));
		}
		hr = S_OK;
		FireOnChanged(DISPID_DTP_CALENDARTODAYDATE);
	}
	return hr;
}

STDMETHODIMP DateTimePicker::get_CalendarTrailingForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(DTM_GETMCCOLOR, MCSC_TRAILINGTEXT, 0));
		if(color != OLECOLOR2COLORREF(properties.calendarTrailingForeColor)) {
			properties.calendarTrailingForeColor = color;
		}
	}

	*pValue = properties.calendarTrailingForeColor;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarTrailingForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARTRAILINGFORECOLOR);
	if(properties.calendarTrailingForeColor != newValue) {
		properties.calendarTrailingForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(DTM_SETMCCOLOR, MCSC_TRAILINGTEXT, OLECOLOR2COLORREF(properties.calendarTrailingForeColor));
		}
		FireOnChanged(DISPID_DTP_CALENDARTRAILINGFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarUseShortestDayNames(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.calendarUseShortestDayNames = ((SendMessage(DTM_GETMCSTYLE, 0, 0) & MCS_SHORTDAYSOFWEEK) == MCS_SHORTDAYSOFWEEK);
	}

	*pValue = BOOL2VARIANTBOOL(properties.calendarUseShortestDayNames);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarUseShortestDayNames(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARUSESHORTESTDAYNAMES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.calendarUseShortestDayNames != b) {
		properties.calendarUseShortestDayNames = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			DWORD style = SendMessage(DTM_GETMCSTYLE, 0, 0);
			if(properties.calendarUseShortestDayNames) {
				style |= MCS_SHORTDAYSOFWEEK;
			} else {
				style &= ~MCS_SHORTDAYSOFWEEK;
			}
			SendMessage(DTM_SETMCSTYLE, 0, style);
		}
		FireOnChanged(DISPID_DTP_CALENDARUSESHORTESTDAYNAMES);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarUseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.calendarUseSystemFont);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarUseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARUSESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.calendarUseSystemFont != b) {
		properties.calendarUseSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyCalendarFont();
		}
		FireOnChanged(DISPID_DTP_CALENDARUSESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CalendarView(ViewConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ViewConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow()) {
		if(IsComctl32Version610OrNewer()) {
			properties.calendarView = static_cast<ViewConstants>(containedCalendar.SendMessage(MCM_GETCURRENTVIEW, 0, 0));
		} else {
			properties.calendarView = vMonth;
		}
	}

	*pValue = properties.calendarView;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CalendarView(ViewConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CALENDARVIEW);
	if(newValue < MCMV_MONTH || newValue > MCMV_MAX) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.calendarView != newValue) {
		properties.calendarView = newValue;
		SetDirty(TRUE);

		if(containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
			containedCalendar.SendMessage(MCM_SETCURRENTVIEW, 0, properties.calendarView);
		}
		FireOnChanged(DISPID_DTP_CALENDARVIEW);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_CheckboxObjectState(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		DATETIMEPICKERINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
		SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
		*pValue = controlInfo.stateCheck;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_CurrentDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SYSTEMTIME systemTime = {0};
		switch(SendMessage(DTM_GETSYSTEMTIME, 0, reinterpret_cast<LPARAM>(&systemTime))) {
			case GDT_VALID:
				SystemTimeToVariantTime(&systemTime, pValue);
				return S_OK;
				break;
			case GDT_NONE:
				SystemTimeToVariantTime(&properties.currentDate, pValue);
				return S_OK;
				break;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::put_CurrentDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CURRENTDATE);
	HRESULT hr = S_OK;
	if(IsWindow()) {
		hr = E_FAIL;
		SYSTEMTIME systemTime = {0};
		DWORD tmp = SendMessage(DTM_GETSYSTEMTIME, 0, reinterpret_cast<LPARAM>(&systemTime));
		if(tmp != GDT_ERROR) {
			properties.dateSelected = (tmp == GDT_VALID);
			VariantTimeToSystemTime(newValue, &systemTime);
			if(!properties.dateSelected) {
				flags.noDateChangeEvent = TRUE;
				SendMessage(DTM_SETSYSTEMTIME, GDT_VALID, reinterpret_cast<LPARAM>(&systemTime));
				flags.noDateChangeEvent = FALSE;
			}
			if(SendMessage(DTM_SETSYSTEMTIME, (properties.dateSelected ? GDT_VALID : GDT_NONE), reinterpret_cast<LPARAM>(&systemTime))) {
				hr = S_OK;
				FireOnChanged(DISPID_DTP_CURRENTDATE);
				FireOnChanged(DISPID_DTP_DATESELECTED);
				FireOnChanged(DISPID_DTP_VALUE);
				Raise_CurrentDateChanged(BOOL2VARIANTBOOL(properties.dateSelected), newValue);
			}
		}
	}
	SendOnDataChange();
	return hr;
}

STDMETHODIMP DateTimePicker::get_CustomDateFormat(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.customDateFormat.Copy();
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_CustomDateFormat(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_DTP_CUSTOMDATEFORMAT);
	if(lstrcmpW(OLE2CW(properties.customDateFormat), OLE2CW(newValue))) {
		properties.customDateFormat = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			/*// DTS_TIMEFORMAT implies DTS_UPDOWN
			DWORD format = GetStyle() & (DTS_SHORTDATEFORMAT | DTS_SHORTDATECENTURYFORMAT | DTS_LONGDATEFORMAT | DTS_TIMEFORMAT);
			if(format != DTS_TIMEFORMAT) {
				format &= ~DTS_UPDOWN;
			}
			if(format == DTS_SHORTDATEFORMAT) {*/
			if(properties.dateFormat == dfCustom) {
				COLE2T converter(properties.customDateFormat);
				LPTSTR pBuffer = converter;
				SendMessage(DTM_SETFORMAT, 0, reinterpret_cast<LPARAM>(pBuffer));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_DTP_CUSTOMDATEFORMAT);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_DateFormat(DateFormatConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DateFormatConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & (DTS_SHORTDATEFORMAT | DTS_SHORTDATECENTURYFORMAT | DTS_LONGDATEFORMAT | DTS_TIMEFORMAT)) {
			case DTS_SHORTDATECENTURYFORMAT:
				properties.dateFormat = dfShortDateWithFourDigitYear;
				break;
			case DTS_LONGDATEFORMAT:
				properties.dateFormat = dfLongDate;
				break;
			case DTS_TIMEFORMAT:
				properties.dateFormat = dfTime;
				break;
			default:
				if(properties.customDateFormat.Length() == 0) {
					properties.dateFormat = dfShortDate;
				} else {
					properties.dateFormat = dfCustom;
				}
				break;
		}
	}

	*pValue = properties.dateFormat;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_DateFormat(DateFormatConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_DATEFORMAT);
	if(properties.dateFormat != newValue) {
		properties.dateFormat = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD upDownStyle = (GetStyle() & DTS_UPDOWN);
			switch(properties.dateFormat) {
				case dfShortDate:
					ModifyStyle(DTS_SHORTDATECENTURYFORMAT | DTS_LONGDATEFORMAT | DTS_TIMEFORMAT, DTS_SHORTDATEFORMAT | upDownStyle);
					SendMessage(DTM_SETFORMAT, 0, NULL);
					break;
				case dfShortDateWithFourDigitYear:
					ModifyStyle(DTS_SHORTDATEFORMAT | DTS_LONGDATEFORMAT | DTS_TIMEFORMAT, DTS_SHORTDATECENTURYFORMAT | upDownStyle);
					SendMessage(DTM_SETFORMAT, 0, NULL);
					break;
				case dfLongDate:
					ModifyStyle(DTS_SHORTDATEFORMAT | DTS_SHORTDATECENTURYFORMAT | DTS_TIMEFORMAT, DTS_LONGDATEFORMAT | upDownStyle);
					SendMessage(DTM_SETFORMAT, 0, NULL);
					break;
				case dfTime:
					// DTS_TIMEFORMAT implies DTS_UPDOWN
					if(upDownStyle) {
						ModifyStyle(DTS_SHORTDATEFORMAT | DTS_SHORTDATECENTURYFORMAT | DTS_LONGDATEFORMAT, DTS_TIMEFORMAT);
						SendMessage(DTM_SETFORMAT, 0, NULL);
					} else {
						properties.customDateFormat = L"";
						RecreateControlWindow();
						FireOnChanged(DISPID_DTP_STYLE);
					}
					break;
				case dfCustom:
					ModifyStyle(DTS_SHORTDATEFORMAT | DTS_SHORTDATECENTURYFORMAT | DTS_LONGDATEFORMAT | DTS_TIMEFORMAT, upDownStyle);
					COLE2T converter(properties.customDateFormat);
					LPTSTR pBuffer = converter;
					SendMessage(DTM_SETFORMAT, 0, reinterpret_cast<LPARAM>(pBuffer));
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_DTP_DATEFORMAT);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_DateSelected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SYSTEMTIME systemTime = {0};
		properties.dateSelected = (SendMessage(DTM_GETSYSTEMTIME, 0, reinterpret_cast<LPARAM>(&systemTime)) == GDT_VALID);
	}

	*pValue = BOOL2VARIANTBOOL(properties.dateSelected);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_DateSelected(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_DATESELECTED);
	BOOL flagSet = properties.allowNullSelection;
	if(!IsInDesignMode() && IsWindow()) {
		flagSet = ((GetStyle() & DTS_SHOWNONE) == DTS_SHOWNONE);
	}
	if(newValue == VARIANT_FALSE && !flagSet) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	properties.dateSelected = VARIANTBOOL2BOOL(newValue);
	SetDirty(TRUE);

	HRESULT hr = S_OK;
	BOOL raiseEvent = FALSE;
	if(IsWindow()) {
		hr = E_FAIL;
		SYSTEMTIME systemTime = {0};
		LRESULT lr = SendMessage(DTM_GETSYSTEMTIME, 0, reinterpret_cast<LPARAM>(&systemTime));
		if(lr != GDT_ERROR) {
			raiseEvent = ((lr == GDT_VALID) != static_cast<BOOL>(properties.dateSelected));
			systemTime = properties.currentDate;
			if(SendMessage(DTM_SETSYSTEMTIME, (properties.dateSelected ? GDT_VALID : GDT_NONE), reinterpret_cast<LPARAM>(&systemTime))) {
				hr = S_OK;
				FireOnChanged(DISPID_DTP_CURRENTDATE);
			} else {
				raiseEvent = FALSE;
			}
		}
	}
	FireOnChanged(DISPID_DTP_DATESELECTED);
	FireOnChanged(DISPID_DTP_VALUE);
	SendOnDataChange();
	if(raiseEvent) {
		DATE date = 0;
		get_CurrentDate(&date);
		Raise_CurrentDateChanged(newValue, date);
	}
	return hr;
}

STDMETHODIMP DateTimePicker::get_DetectDoubleClicks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_DetectDoubleClicks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_DETECTDOUBLECLICKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDoubleClicks != b) {
		properties.detectDoubleClicks = b;
		SetDirty(TRUE);

		FireOnChanged(DISPID_DTP_DETECTDOUBLECLICKS);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		/*if(SendMessage(DTM_GETMCSTYLE, 0, 0) & MCS_DAYSTATE) {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents & ~deGetBoldDates);
		} else {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents | deGetBoldDates);
		}*/
		if(GetStyle() & DTS_APPCANPARSE) {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents & ~deParseUserInput);
		} else {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents | deParseUserInput);
		}
	}
	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deMouseEvents) != (newValue & deMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}
		if((properties.disabledEvents & deCalendarMouseEvents) != (newValue & deCalendarMouseEvents)) {
			if(containedCalendar.IsWindow()) {
				if(newValue & deCalendarMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = containedCalendar;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
					dateUnderMouse = 0;
				}
			}
		}
		/*if((properties.disabledEvents & deGetBoldDates) != (newValue & deGetBoldDates)) {
			if(IsWindow() && IsComctl32Version610OrNewer()) {
				DWORD style = SendMessage(DTM_GETMCSTYLE, 0, 0);
				if(newValue & deGetBoldDates) {
					style &= ~MCS_DAYSTATE;
				} else {
					style |= MCS_DAYSTATE;
				}
				SendMessage(DTM_SETMCSTYLE, 0, style);
			}
		}*/
		if((properties.disabledEvents & deParseUserInput) != (newValue & deParseUserInput)) {
			if(IsWindow()) {
				if(newValue & deParseUserInput) {
					ModifyStyle(DTS_APPCANPARSE, 0);
				} else {
					ModifyStyle(0, DTS_APPCANPARSE);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_DTP_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_DragDropDownTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragDropDownTime;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_DragDropDownTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DTP_DRAGDROPDOWNTIME);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragDropDownTime != newValue) {
		properties.dragDropDownTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_DRAGDROPDOWNTIME);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_DropDownAlignment(DropDownAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DropDownAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch(GetStyle() & DTS_RIGHTALIGN) {
			case DTS_RIGHTALIGN:
				properties.dropDownAlignment = ddaRight;
				break;
			case 0:
				properties.dropDownAlignment = ddaLeft;
				break;
		}
	}

	*pValue = properties.dropDownAlignment;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_DropDownAlignment(DropDownAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_DROPDOWNALIGNMENT);
	if(properties.dropDownAlignment != newValue) {
		properties.dropDownAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.dropDownAlignment) {
				case ddaLeft:
					ModifyStyle(DTS_RIGHTALIGN, 0);
					break;
				case ddaRight:
					ModifyStyle(0, DTS_RIGHTALIGN);
					break;
			}
		}
		FireOnChanged(DISPID_DTP_DROPDOWNALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_DropDownButtonObjectState(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		DATETIMEPICKERINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
		SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
		*pValue = controlInfo.stateButton;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_DTP_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_DTP_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DTP_FONT);
	return S_OK;
}

STDMETHODIMP DateTimePicker::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_DTP_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DTP_FONT);
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_DTP_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_hWndCalendar(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(DTM_GETMONTHCAL, 0, 0)));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_hWndEdit(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(IsComctl32Version610OrNewer()) {
			DATETIMEPICKERINFO controlInfo = {0};
			controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
			SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
			*pValue = HandleToLong(controlInfo.hwndEdit);
		} else {
			*pValue = HandleToLong(FindWindowEx(*this, NULL, WC_EDIT, NULL));
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_hWndUpDown(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(IsComctl32Version610OrNewer()) {
			DATETIMEPICKERINFO controlInfo = {0};
			controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
			SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
			*pValue = HandleToLong(controlInfo.hwndUD);
		} else {
			*pValue = HandleToLong(static_cast<HWND>(GetDlgItem(1000)));
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::get_IMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.IMEMode;
	} else {
		if((GetFocus() == *this) && (GetEffectiveIMEMode() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(*this);
			if((ime != imeInherit) && (properties.IMEMode != imeInherit)) {
				properties.IMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode();
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == *this) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(*this, ime);
				}
			}
		}
		FireOnChanged(DISPID_DTP_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_MaxDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		DWORD setBoundaries = SendMessage(DTM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
		if(!(setBoundaries & GDTR_MAX)) {
			boundaryTimes[1].wYear = 9999;
			boundaryTimes[1].wMonth = 12;
			boundaryTimes[1].wDay = 31;
			/*boundaryTimes[1].wHour = 23;
			boundaryTimes[1].wMinute = 59;
			boundaryTimes[1].wSecond = 59;
			//boundaryTimes[1].wMilliseconds = 999;*/
		}
		SystemTimeToVariantTime(&boundaryTimes[1], &properties.maxDate);
	}

	*pValue = properties.maxDate;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_MaxDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_DTP_MAXDATE);
	if(properties.maxDate != newValue) {
		properties.maxDate = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SYSTEMTIME boundaryTimes[2];
			ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
			// According to the documentation you can set min and max independently, but this does not seem to work.
			SendMessage(DTM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
			DWORD setBoundaries = SendMessage(DTM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
			if(!(setBoundaries & GDTR_MIN)) {
				boundaryTimes[0].wYear = 1752;
				boundaryTimes[0].wMonth = 9;
				boundaryTimes[0].wDay = 14;
			}
			VariantTimeToSystemTime(properties.maxDate, &boundaryTimes[1]);
			SendMessage(DTM_SETRANGE, GDTR_MIN | GDTR_MAX, reinterpret_cast<LPARAM>(&boundaryTimes));
		}
		FireOnChanged(DISPID_DTP_MAXDATE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_MinDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		DWORD setBoundaries = SendMessage(DTM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
		if(!(setBoundaries & GDTR_MIN)) {
			boundaryTimes[0].wYear = 1752;
			boundaryTimes[0].wMonth = 9;
			boundaryTimes[0].wDay = 14;
		}
		SystemTimeToVariantTime(&boundaryTimes[0], &properties.minDate);
	}

	*pValue = properties.minDate;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_MinDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_DTP_MINDATE);
	if(properties.minDate != newValue) {
		properties.minDate = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SYSTEMTIME boundaryTimes[2];
			ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
			// According to the documentation you can set min and max independently, but this does not seem to work.
			DWORD setBoundaries = SendMessage(DTM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
			if(!(setBoundaries & GDTR_MAX)) {
				boundaryTimes[1].wYear = 9999;
				boundaryTimes[1].wMonth = 12;
				boundaryTimes[1].wDay = 31;
				/*boundaryTimes[1].wHour = 23;
				boundaryTimes[1].wMinute = 59;
				boundaryTimes[1].wSecond = 59;
				//boundaryTimes[1].wMilliseconds = 999;*/
			}
			VariantTimeToSystemTime(properties.minDate, &boundaryTimes[0]);
			SendMessage(DTM_SETRANGE, GDTR_MIN | GDTR_MAX, reinterpret_cast<LPARAM>(&boundaryTimes));
		}
		FireOnChanged(DISPID_DTP_MINDATE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_DTP_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DTP_MOUSEICON);
	return S_OK;
}

STDMETHODIMP DateTimePicker::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_DTP_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_DTP_MOUSEICON);
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
			if(containedCalendar.IsWindow()) {
				if(properties.registerForOLEDragDrop) {
					ATLVERIFY(RegisterDragDrop(containedCalendar, static_cast<IDropTarget*>(this)) == S_OK);
				} else {
					ATLVERIFY(RevokeDragDrop(containedCalendar) == S_OK);
				}
			}
		}
		FireOnChanged(DISPID_DTP_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = GetExStyle();
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeft & rtlLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
			}
			if(properties.rightToLeft & rtlText) {
				ModifyStyleEx(0, WS_EX_RTLREADING);
			} else {
				ModifyStyleEx(WS_EX_RTLREADING, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_DTP_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_StartOfWeek(StartOfWeekConstants* pValue)
{
	ATLASSERT_POINTER(pValue, StartOfWeekConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedCalendar.IsWindow()) {
		DWORD tmp = containedCalendar.SendMessage(MCM_GETFIRSTDAYOFWEEK, 0, 0);
		if(HIWORD(tmp)) {
			properties.startOfWeek = static_cast<StartOfWeekConstants>(LOWORD(tmp));
		} else {
			properties.startOfWeek = sowUseLocale;
		}
	}

	*pValue = properties.startOfWeek;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_StartOfWeek(StartOfWeekConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_STARTOFWEEK);
	if(newValue < sowUseLocale || newValue > sowSunday) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.startOfWeek != newValue) {
		properties.startOfWeek = newValue;
		SetDirty(TRUE);

		if(containedCalendar.IsWindow()) {
			containedCalendar.SendMessage(MCM_SETFIRSTDAYOFWEEK, 0, properties.startOfWeek);
		}
		FireOnChanged(DISPID_DTP_STARTOFWEEK);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_Style(StyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, StyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.style = (GetStyle() & DTS_UPDOWN ? sUpDown : sDropDown);
	}

	*pValue = properties.style;
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_Style(StyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_DTP_STYLE);
	BOOL displayingTime = (properties.dateFormat == dfTime);
	if(!IsInDesignMode() && IsWindow()) {
		displayingTime = ((GetStyle() & DTS_TIMEFORMAT) == DTS_TIMEFORMAT);
	}
	if(newValue == sDropDown && displayingTime) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.style != newValue) {
		properties.style = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.style) {
				case sUpDown:
					if(!(GetStyle() & DTS_UPDOWN)) {
						RecreateControlWindow();
					}
					break;
				case sDropDown:
					if(GetStyle() & DTS_UPDOWN) {
						RecreateControlWindow();
					}
					break;
			}
		}
		FireOnChanged(DISPID_DTP_STYLE);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP DateTimePicker::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_DTP_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_DTP_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::get_Value(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr;

	VariantClear(pValue);
	if(GetStyle() & DTS_SHOWNONE) {
		// NULL date is possible
		VARIANT_BOOL b = VARIANT_FALSE;
		hr = get_DateSelected(&b);
		if(SUCCEEDED(hr)) {
			if(b != VARIANT_FALSE) {
				// a date is selected - return it
				hr = get_CurrentDate(&pValue->date);
				if(SUCCEEDED(hr)) {
					pValue->vt = VT_DATE;
				}
			} else {
				// no date is selected - return NULL
				pValue->vt = VT_NULL;
			}
		}
	} else {
		// NULL date is not possible, so return the date
		hr = get_CurrentDate(&pValue->date);
		if(SUCCEEDED(hr)) {
			pValue->vt = VT_DATE;
		}
	}
	return hr;
}

STDMETHODIMP DateTimePicker::put_Value(VARIANT newValue)
{
	PUTPROPPROLOG(DISPID_DTP_VALUE);
	HRESULT hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	if(GetStyle() & DTS_SHOWNONE) {
		// NULL selection is possible
		switch(newValue.vt) {
			case VT_DATE:
				flags.noDateChangeEvent = TRUE;
				hr = put_DateSelected(VARIANT_TRUE);
				flags.noDateChangeEvent = FALSE;
				if(SUCCEEDED(hr)) {
					hr = put_CurrentDate(newValue.date);
				}
				break;
			case VT_EMPTY:
			case VT_NULL:
				hr = put_DateSelected(VARIANT_FALSE);
				break;
			default:
				VARIANT v;
				VariantInit(&v);
				hr = VariantChangeType(&v, &newValue, 0, VT_DATE);
				if(SUCCEEDED(hr)) {
					flags.noDateChangeEvent = TRUE;
					hr = put_DateSelected(VARIANT_TRUE);
					flags.noDateChangeEvent = FALSE;
				}
				if(SUCCEEDED(hr)) {
					hr = put_CurrentDate(v.date);
				}
				VariantClear(&v);
				break;
		}
	} else {
		// NULL selection is not possible
		VARIANT v;
		VariantInit(&v);
		hr = VariantChangeType(&v, &newValue, 0, VT_DATE);
		if(SUCCEEDED(hr)) {
			hr = put_CurrentDate(v.date);
		}
		VariantClear(&v);
	}
	if(SUCCEEDED(hr)) {
		FireOnChanged(DISPID_DTP_VALUE);
	}
	SendOnDataChange();
	return hr;
}

STDMETHODIMP DateTimePicker::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP DateTimePicker::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP DateTimePicker::CalendarGetMaxTodayWidth(OLE_XSIZE_PIXELS* pWidth)
{
	ATLASSERT_POINTER(pWidth, OLE_XSIZE_PIXELS);
	if(!pWidth) {
		return E_POINTER;
	}

	if(containedCalendar.IsWindow()) {
		*pWidth = containedCalendar.SendMessage(MCM_GETMAXTODAYWIDTH);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::CalendarGetRectangle(ControlPartConstants controlPart, LONG columnIndex/* = 0*/, LONG rowIndex/* = 0*/, OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(containedCalendar.IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = controlPart;
		gridInfo.dwFlags = MCGIF_RECT;
		gridInfo.iCol = columnIndex;
		gridInfo.iRow = rowIndex;
		if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			if(pLeft) {
				*pLeft = gridInfo.rc.left;
			}
			if(pTop) {
				*pTop = gridInfo.rc.top;
			}
			if(pRight) {
				*pRight = gridInfo.rc.right;
			}
			if(pBottom) {
				*pBottom = gridInfo.rc.bottom;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::CloseDropDownWindow(void)
{
	if(IsWindow()) {
		if(IsComctl32Version610OrNewer()) {
			SendMessage(DTM_CLOSEMONTHCAL, 0, 0);
		} else if(!(GetStyle() & DTS_UPDOWN) && SendMessage(DTM_GETMONTHCAL, 0, 0)) {
			flags.ignoreEscapeKeyDownAndUp = TRUE;
			flags.ignoreEscapeChar = 2;
			PostMessage(WM_KEYDOWN, VK_ESCAPE, 0);
			PostMessage(WM_KEYUP, VK_ESCAPE, 0);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP DateTimePicker::GetCheckboxRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(IsWindow() && IsComctl32Version610OrNewer()) {
		DATETIMEPICKERINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
		SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
		if(pLeft) {
			*pLeft = controlInfo.rcCheck.left;
		}
		if(pTop) {
			*pTop = controlInfo.rcCheck.top;
		}
		if(pRight) {
			*pRight = controlInfo.rcCheck.right;
		}
		if(pBottom) {
			*pBottom = controlInfo.rcCheck.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::GetDropDownButtonRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(IsWindow() && IsComctl32Version610OrNewer()) {
		DATETIMEPICKERINFO controlInfo = {0};
		controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
		SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
		if(pLeft) {
			*pLeft = controlInfo.rcButton.left;
		}
		if(pTop) {
			*pTop = controlInfo.rcButton.top;
		}
		if(pRight) {
			*pRight = controlInfo.rcButton.right;
		}
		if(pBottom) {
			*pBottom = controlInfo.rcButton.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::GetIdealSize(OLE_XSIZE_PIXELS* pWidth/* = NULL*/, OLE_YSIZE_PIXELS* pHeight/* = NULL*/)
{
	if(IsWindow() && IsComctl32Version610OrNewer()) {
		SIZE idealSize = {0};
		if(SendMessage(DTM_GETIDEALSIZE, 0, reinterpret_cast<LPARAM>(&idealSize))) {
			RECT windowRectangle = {0, 0, idealSize.cx, idealSize.cy};
			AdjustWindowRectEx(&windowRectangle, GetStyle(), FALSE, GetExStyle());
			idealSize.cx = windowRectangle.right - windowRectangle.left;
			idealSize.cy = windowRectangle.bottom - windowRectangle.top;
			if(pWidth) {
				*pWidth = idealSize.cx;
			}
			if(pHeight) {
				*pHeight = idealSize.cy;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::OpenDropDownWindow(void)
{
	if(IsWindow()) {
		if(!(GetStyle() & DTS_UPDOWN) && !SendMessage(DTM_GETMONTHCAL, 0, 0)) {
			// the F4 approach has problems with callback fields
			/*DefWindowProc(WM_KEYDOWN, VK_F4, 0);
			DefWindowProc(WM_KEYUP, VK_F4, 0);*/
			RECT buttonRectangle = {0};
			GetClientRect(&buttonRectangle);
			buttonRectangle.left = buttonRectangle.right - GetSystemMetrics(SM_CXVSCROLL);
			DefWindowProc(WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(buttonRectangle.left, buttonRectangle.top));
			DefWindowProc(WM_LBUTTONUP, MK_LBUTTON, MAKELPARAM(buttonRectangle.left, buttonRectangle.top));
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP DateTimePicker::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP DateTimePicker::SetBoldDates(LPSAFEARRAY* ppStates, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(ppStates, LPSAFEARRAY);
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!ppStates || !pSucceeded) {
		return E_POINTER;
	}
	LPSAFEARRAY pStates = *ppStates;
	ATLASSERT_POINTER(pStates, SAFEARRAY);
	if(!pStates) {
		return E_INVALIDARG;
	}

	*pSucceeded = VARIANT_FALSE;

	VARTYPE vt = VT_EMPTY;
	ATLVERIFY(SUCCEEDED(SafeArrayGetVartype(pStates, &vt)));
	if(vt != VT_BOOL) {
		return E_INVALIDARG;
	}

	if(containedCalendar.IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		int numberOfMonths = containedCalendar.SendMessage(MCM_GETMONTHRANGE, GMR_DAYSTATE, reinterpret_cast<LPARAM>(&boundaryTimes));
		LONG l;
		ATLVERIFY(SUCCEEDED(SafeArrayGetLBound(pStates, 1, &l)));
		LONG u;
		ATLVERIFY(SUCCEEDED(SafeArrayGetUBound(pStates, 1, &u)));
		if(u < l || u - l + 1 < numberOfMonths * 31) {
			return E_INVALIDARG;
		}

		LPMONTHDAYSTATE pMonthStates = static_cast<LPMONTHDAYSTATE>(HeapAlloc(GetProcessHeap(), 0, numberOfMonths * sizeof(MONTHDAYSTATE)));
		if(!pMonthStates) {
			return E_OUTOFMEMORY;
		}
		ZeroMemory(pMonthStates, numberOfMonths * sizeof(MONTHDAYSTATE));

		VARIANT_BOOL b;
		LONG overallIndex = l;
		for(int i = 0; i < numberOfMonths; ++i) {
			for(int day = 0; day < 31 && overallIndex <= u; ++day, ++overallIndex) {
				ATLVERIFY(SUCCEEDED(SafeArrayGetElement(pStates, &overallIndex, &b)));
				if(b != VARIANT_FALSE) {
					pMonthStates[i] |= (0x00000001 << day);
				}
			}
		}

		*pSucceeded = BOOL2VARIANTBOOL(containedCalendar.SendMessage(MCM_SETDAYSTATE, numberOfMonths, reinterpret_cast<LPARAM>(pMonthStates)));

		HeapFree(GetProcessHeap(), 0, pMonthStates);
		return S_OK;
	}
	return E_FAIL;
}


LRESULT DateTimePicker::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(flags.ignoreEscapeChar == 0 || keyAscii != VK_ESCAPE) {
			if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
				// the client may have changed the key code (actually it can be changed to 0 only)
				wParam = keyAscii;
				if(wParam == 0) {
					wasHandled = TRUE;
				}
			}
		} else {
			flags.ignoreEscapeChar--;
		}
	}
	return 0;
}

LRESULT DateTimePicker::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y, &showDefaultMenu);
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}
	return 0;
}

LRESULT DateTimePicker::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// we want to receive WM_*BUTTONDBLCLK messages
	DWORD style = GetClassLong(*this, GCL_STYLE);
	style |= CS_DBLCLKS;
	SetClassLong(*this, GCL_STYLE, style);

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT DateTimePicker::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT DateTimePicker::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if((ime != imeNoControl) && (GetFocus() == *this)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(*this, ime);
	} else if((ime != imeNoControl) && (GetFocus() == containedEdit)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(containedEdit, ime);
	}
	return lr;
}

LRESULT DateTimePicker::OnKeyDown(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(index != 0 || !flags.ignoreEscapeKeyDownAndUp || keyCode != VK_ESCAPE) {
			if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
				// the client may have changed the key code
				wParam = keyCode;
				if(wParam == 0) {
					return 0;
				}
			}
		}
	}
	if(index == 2) {
		return containedEdit.DefWindowProc(message, wParam, lParam);
	} else {
		return DefWindowProc(message, wParam, lParam);
	}
}

LRESULT DateTimePicker::OnKeyUp(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(index != 0 || !flags.ignoreEscapeKeyDownAndUp || keyCode != VK_ESCAPE) {
			if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
				// the client may have changed the key code
				wParam = keyCode;
				if(wParam == 0) {
					return 0;
				}
			}
		}
	}
	if(index == 2) {
		return containedEdit.DefWindowProc(message, wParam, lParam);
	} else {
		if(index == 0 && flags.ignoreEscapeKeyDownAndUp && wParam == VK_ESCAPE) {
			flags.ignoreEscapeKeyDownAndUp = FALSE;
		}
		return DefWindowProc(message, wParam, lParam);
	}
}

LRESULT DateTimePicker::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<DateTimePicker>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT DateTimePicker::OnLButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(index == 0 && flags.ignoreDblClick) {
		flags.ignoreDblClick = FALSE;
	} else {
		if(!properties.detectDoubleClicks) {
			if(index == 2) {
				return containedEdit.SendMessage(WM_LBUTTONDOWN, wParam, lParam);
			} else {
				return SendMessage(WM_LBUTTONDOWN, wParam, lParam);
			}
		}

		wasHandled = FALSE;

		if(!(properties.disabledEvents & deClickEvents)) {
			POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
			if(index == 2) {
				containedEdit.MapWindowPoints(*this, &pt, 1);
			}

			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(wParam, button, shift);
			button = 1/*MouseButtonConstants.vbLeftButton*/;
			Raise_DblClick(button, shift, pt.x, pt.y);
		}
	}

	return 0;
}

LRESULT DateTimePicker::OnLButtonDown(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(index == 0 && !m_bUIActive && !RunTimeHelper::IsCommCtrl6()) {
		if(!flags.ignoreDblClick) {
			flags.applyDropDownWorkaround = TRUE;
			bufferedLButtonDown.hwnd = *this;
			bufferedLButtonDown.message = message;
			bufferedLButtonDown.wParam = wParam;
			bufferedLButtonDown.lParam = lParam;
			bufferedLButtonDown.time = GetMessageTime();
			DWORD pos = GetMessagePos();
			bufferedLButtonDown.pt.x = GET_X_LPARAM(pos);
			bufferedLButtonDown.pt.y = GET_Y_LPARAM(pos);
		}
		return 0;
	}

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if(index == 2) {
		containedEdit.MapWindowPoints(*this, &pt, 1);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(index, button, shift, pt.x, pt.y);

	return 0;
}

LRESULT DateTimePicker::OnLButtonUp(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	static BOOL recursiveCall = FALSE;

	if(index == 0 && flags.applyDropDownWorkaround) {
		// fake two mouse-clicks
		flags.applyDropDownWorkaround = FALSE;
		INPUT input[2];
		input[0].type = INPUT_MOUSE;
		input[0].mi.dx = 0;
		input[0].mi.dy = 0;
		input[0].mi.mouseData = 0;
		input[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
		input[0].mi.time = bufferedLButtonDown.time;
		input[0].mi.dwExtraInfo = 0;
		input[1].type = INPUT_MOUSE;
		input[1].mi.dx = 0;
		input[1].mi.dy = 0;
		input[1].mi.mouseData = 0;
		input[1].mi.dwFlags = MOUSEEVENTF_LEFTUP | MOUSEEVENTF_MOVE;
		input[1].mi.time = GetMessageTime();
		input[1].mi.dwExtraInfo = 0;
		ReplyMessage(DefWindowProc(message, wParam, lParam));
		// the first click will cause a double click, which we will ignore
		flags.ignoreDblClick = TRUE;
		recursiveCall = TRUE;
		SendInput(2, input, sizeof(INPUT));
		// this causes the real click
		SendInput(2, input, sizeof(INPUT));
		return 0;
	}
	wasHandled = FALSE;

	if(recursiveCall) {
		recursiveCall = FALSE;
	} else {
		POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(index == 2) {
			containedEdit.MapWindowPoints(*this, &pt, 1);
		}

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseUp(index, button, shift, pt.x, pt.y);
	}
	return 0;
}

LRESULT DateTimePicker::OnMButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		if(index == 2) {
			return containedEdit.SendMessage(WM_MBUTTONDOWN, wParam, lParam);
		} else {
			return SendMessage(WM_MBUTTONDOWN, wParam, lParam);
		}
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(index == 2) {
			containedEdit.MapWindowPoints(*this, &pt, 1);
		}

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, pt.x, pt.y);
	}

	return 0;
}

LRESULT DateTimePicker::OnMButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if(index == 2) {
		containedEdit.MapWindowPoints(*this, &pt, 1);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(index, button, shift, pt.x, pt.y);

	return 0;
}

LRESULT DateTimePicker::OnMButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if(index == 2) {
		containedEdit.MapWindowPoints(*this, &pt, 1);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(index, button, shift, pt.x, pt.y);

	return 0;
}

LRESULT DateTimePicker::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT DateTimePicker::OnMouseHover(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if(index == 2) {
		containedEdit.MapWindowPoints(*this, &pt, 1);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, pt.x, pt.y);

	return 0;
}

LRESULT DateTimePicker::OnMouseLeave(LONG /*index*/, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT DateTimePicker::OnMouseMove(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(index == 0 && flags.applyDropDownWorkaround)) {
		if(!(properties.disabledEvents & deMouseEvents)) {
			POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
			if(index == 2) {
				containedEdit.MapWindowPoints(*this, &pt, 1);
			}

			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(wParam, button, shift);
			Raise_MouseMove(index, button, shift, pt.x, pt.y);
		}
	}

	return 0;
}

LRESULT DateTimePicker::OnMouseWheel(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_MouseWheel(index, button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
	}

	return 0;
}

LRESULT DateTimePicker::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT DateTimePicker::OnParentNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = S_OK;
	wasHandled = FALSE;

	switch(wParam) {
		case WM_CREATE: {
			lr = DefWindowProc(message, wParam, lParam);
			wasHandled = TRUE;

			LPTSTR pBuffer = new TCHAR[MAX_PATH + 1];
			GetClassName(reinterpret_cast<HWND>(lParam), pBuffer, MAX_PATH);
			if(lstrcmp(pBuffer, reinterpret_cast<LPCTSTR>(&WC_EDIT)) == 0) {
				Raise_CreatedEditControlWindow(static_cast<LONG>(lParam));
			}
			delete[] pBuffer;
			break;
		}
	}
	return lr;
}

LRESULT DateTimePicker::OnRButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		if(index == 2) {
			return containedEdit.SendMessage(WM_RBUTTONDOWN, wParam, lParam);
		} else {
			return SendMessage(WM_RBUTTONDOWN, wParam, lParam);
		}
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(index == 2) {
			containedEdit.MapWindowPoints(*this, &pt, 1);
		}

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, pt.x, pt.y);
	}

	return 0;
}

LRESULT DateTimePicker::OnRButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if(index == 2) {
		containedEdit.MapWindowPoints(*this, &pt, 1);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(index, button, shift, pt.x, pt.y);

	return 0;
}

LRESULT DateTimePicker::OnRButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if(index == 2) {
		containedEdit.MapWindowPoints(*this, &pt, 1);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(index, button, shift, pt.x, pt.y);

	return 0;
}

LRESULT DateTimePicker::OnSetCursor(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	WTL::CRect clientArea;
	if(index == 2) {
		containedEdit.GetClientRect(&clientArea);
		containedEdit.ClientToScreen(&clientArea);
	} else {
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
	}
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(index == 2) {
			if(WindowFromPoint(mousePosition) == containedEdit) {
				setCursor = TRUE;
			}
		} else {
			if(WindowFromPoint(mousePosition) == *this) {
				setCursor = TRUE;
			}
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT DateTimePicker::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<DateTimePicker>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<DateTimePicker>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT DateTimePicker::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DTP_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_FONT);
	}

	return lr;
}

LRESULT DateTimePicker::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT DateTimePicker::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT DateTimePicker::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT DateTimePicker::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				BOOL dummy = TRUE;
				OnTimer(WM_TIMER, timers.ID_SETCUSTOMDATEFORMAT, 0, dummy);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		case timers.ID_SETCUSTOMDATEFORMAT:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_SETCUSTOMDATEFORMAT);
				if(properties.dateFormat == dfCustom) {
					// we won't receive DTN_FORMATQUERY notifications if we do this in SendConfigurationMessages
					COLE2T converter = properties.customDateFormat;
					LPTSTR pBuffer = converter;
					SendMessage(DTM_SETFORMAT, 0, reinterpret_cast<LPARAM>(pBuffer));
				}
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		case timers.ID_DRAGDROPDOWN:
			KillTimer(timers.ID_DRAGDROPDOWN);
			dragDropStatus.dropDownTimerIsRunning = FALSE;
			OpenDropDownWindow();
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT DateTimePicker::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT DateTimePicker::OnXButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		switch(index) {
			case 0:
				Raise_XDblClick(button, shift, mousePosition.x, mousePosition.y);
				break;
			/*case 2:
				containedEdit.MapWindowPoints(*this, &mousePosition, 1);
				Raise_XDblClick(button, shift, mousePosition.x, mousePosition.y);
				break;*/
		}
	}

	return 0;
}

LRESULT DateTimePicker::OnXButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseDown(index, button, shift, mousePosition.x, mousePosition.y);
			break;
		/*case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseDown(index, button, shift, mousePosition.x, mousePosition.y);
			break;*/
	}

	return 0;
}

LRESULT DateTimePicker::OnXButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	switch(index) {
		case 0:
			Raise_MouseUp(index, button, shift, mousePosition.x, mousePosition.y);
			break;
		/*case 2:
			containedEdit.MapWindowPoints(*this, &mousePosition, 1);
			Raise_MouseUp(index, button, shift, mousePosition.x, mousePosition.y);
			break;*/
	}

	return 0;
}

LRESULT DateTimePicker::OnReflectedNotifyFormat(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == NF_QUERY) {
		#ifdef UNICODE
			return NFR_UNICODE;
		#else
			return NFR_ANSI;
		#endif
	}
	return 0;
}

LRESULT DateTimePicker::OnGetMCStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(IsComctl32Version610OrNewer()) {
		return DefWindowProc(message, wParam, lParam);
	} else {
		HWND hWndCalendar = reinterpret_cast<HWND>(SendMessage(DTM_GETMONTHCAL, 0, 0));
		if(hWndCalendar) {
			properties.calendarStyle = CWindow(hWndCalendar).GetStyle() & MCS_MASK;
		}
		return properties.calendarStyle;
	}
}

LRESULT DateTimePicker::OnSetFormat(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		if(lParam) {
			if(message == DTM_SETFORMATA) {
				properties.customDateFormat = A2BSTR(reinterpret_cast<LPSTR>(lParam));
			} else {
				properties.customDateFormat = W2BSTR(reinterpret_cast<LPWSTR>(lParam));
			}
			properties.dateFormat = dfCustom;
		} else {
			// DTS_TIMEFORMAT implies DTS_UPDOWN
			DWORD format = GetStyle() & (DTS_SHORTDATEFORMAT | DTS_SHORTDATECENTURYFORMAT | DTS_LONGDATEFORMAT | DTS_TIMEFORMAT);
			if(format != DTS_TIMEFORMAT) {
				format &= ~DTS_UPDOWN;
			}
			switch(format) {
				case DTS_SHORTDATECENTURYFORMAT:
					properties.dateFormat = dfShortDateWithFourDigitYear;
					break;
				case DTS_LONGDATEFORMAT:
					properties.dateFormat = dfLongDate;
					break;
				case DTS_TIMEFORMAT:
					properties.dateFormat = dfTime;
					break;
				default:
					properties.dateFormat = dfShortDate;
					break;
			}
		}
		FireOnChanged(DISPID_DTP_CUSTOMDATEFORMAT);
		FireOnChanged(DISPID_DTP_DATEFORMAT);
	}
	return lr;
}

LRESULT DateTimePicker::OnSetMCFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DTP_CALENDARFONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.calendarFont.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new calendarFont.pFontDisp object
		if(!properties.calendarFont.owningFontResource) {
			properties.calendarFont.currentFont.Detach();
		}
		properties.calendarFont.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.calendarFont.owningFontResource = FALSE;
		properties.calendarUseSystemFont = FALSE;
		properties.calendarFont.StopWatching();

		if(properties.calendarFont.pFontDisp) {
			properties.calendarFont.pFontDisp->Release();
			properties.calendarFont.pFontDisp = NULL;
		}
		if(!properties.calendarFont.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.calendarFont.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.calendarFont.pFontDisp));
			}
		}
		properties.calendarFont.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_DTP_CALENDARFONT);
	}

	return lr;
}

LRESULT DateTimePicker::OnSetMCStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(IsComctl32Version610OrNewer()) {
		return DefWindowProc(message, wParam, lParam);
	} else {
		DWORD oldStyle = properties.calendarStyle;
		properties.calendarStyle = lParam & MCS_MASK;
		HWND hWndCalendar = reinterpret_cast<HWND>(SendMessage(DTM_GETMONTHCAL, 0, 0));
		if(hWndCalendar) {
			CWindow(hWndCalendar).ModifyStyle(MCS_MASK, properties.calendarStyle);
		}
		return oldStyle;
	}
}

LRESULT DateTimePicker::OnSetSystemTime(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_DTP_CURRENTDATE) == S_FALSE || FireOnRequestEdit(DISPID_DTP_DATESELECTED) == S_FALSE || FireOnRequestEdit(DISPID_DTP_VALUE) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		if(wParam == GDT_VALID) {
			properties.currentDate = *reinterpret_cast<LPSYSTEMTIME>(lParam);
		}
	}
	FireOnChanged(DISPID_DTP_CURRENTDATE);
	FireOnChanged(DISPID_DTP_DATESELECTED);
	FireOnChanged(DISPID_DTP_VALUE);
	SendOnDataChange();
	return lr;
}


LRESULT DateTimePicker::OnCalendarChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deCalendarKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_CalendarKeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT DateTimePicker::OnCalendarContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		containedCalendar.ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_CalendarContextMenu(button, shift, mousePosition.x, mousePosition.y, &showDefaultMenu);
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT DateTimePicker::OnCalendarDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!flags.applyDropDownWorkaround) {
		Raise_DestroyedCalendarControlWindow(HandleToLong(static_cast<HWND>(containedCalendar)));
	}
	return 0;
}

LRESULT DateTimePicker::OnCalendarKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deCalendarKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_CalendarKeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedCalendar.DefWindowProc(message, wParam, lParam);
}

LRESULT DateTimePicker::OnCalendarKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deCalendarKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_CalendarKeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedCalendar.DefWindowProc(message, wParam, lParam);
}

LRESULT DateTimePicker::OnCalendarLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_CalendarMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DateTimePicker::OnCalendarLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_CalendarMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DateTimePicker::OnCalendarMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_CalendarMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DateTimePicker::OnCalendarMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_CalendarMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DateTimePicker::OnCalendarMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_CalendarMouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DateTimePicker::OnCalendarMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_CalendarMouseLeave(button, shift, mouseStatus_Calendar.lastPosition.x, mouseStatus_Calendar.lastPosition.y);

	return 0;
}

LRESULT DateTimePicker::OnCalendarMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deCalendarMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_CalendarMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT DateTimePicker::OnCalendarMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deCalendarMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_CalendarMouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else {
		if(!mouseStatus_Calendar.enteredControl) {
			mouseStatus_Calendar.EnterControl();
		}
	}

	return 0;
}

LRESULT DateTimePicker::OnCalendarRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_CalendarMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DateTimePicker::OnCalendarRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_CalendarMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT DateTimePicker::OnCalendarXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	Raise_CalendarMouseDown(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT DateTimePicker::OnCalendarXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	Raise_CalendarMouseUp(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT DateTimePicker::OnEditDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedEditControlWindow(HandleToLong(static_cast<HWND>(containedEdit)));
	return 0;
}


LRESULT DateTimePicker::OnSetFocusNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(*this);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(*this, ime);
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT DateTimePicker::OnCloseUpNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(!flags.applyDropDownWorkaround) {
		Raise_CalendarCloseUp();
	}
	return 0;
}

LRESULT DateTimePicker::OnDateTimeChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMDATETIMECHANGE pDetails = reinterpret_cast<LPNMDATETIMECHANGE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMDATETIMECHANGE);
	if(!pDetails) {
		return 0;
	}

	if(pDetails->dwFlags == GDT_VALID) {
		properties.currentDate = pDetails->st;
		properties.dateSelected = TRUE;
	} else {
		properties.dateSelected = FALSE;
	}
	FireOnChanged(DISPID_DTP_CURRENTDATE);
	FireOnChanged(DISPID_DTP_DATESELECTED);
	FireOnChanged(DISPID_DTP_VALUE);
	SendOnDataChange();

	DATE date = 0;
	if(pDetails->dwFlags == GDT_VALID) {
		SystemTimeToVariantTime(&pDetails->st, &date);
	}
	Raise_CurrentDateChanged(BOOL2VARIANTBOOL(pDetails->dwFlags == GDT_VALID), date);
	return 0;
}

LRESULT DateTimePicker::OnDropDownNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(flags.applyDropDownWorkaround) {
		// It may happen that the work-around was not necessary at all. But we don't know this in advance.
		CloseDropDownWindow();
	} else {
		Raise_CalendarDropDown();

		if(GetAsyncKeyState(VK_LBUTTON) & 0x8000) {
			/* HACK: We capture mouse input on mouse-down and this interferes with drop-down handling of
							 SysDateTimePick32. */
			if(mouseStatus.IsClickCandidate(1/*MouseButtonConstants.vbLeftButton*/)) {
				if(GetCapture() == *this) {
					ReleaseCapture();
				}
				mouseStatus.RemoveClickCandidate(1/*MouseButtonConstants.vbLeftButton*/);
			}
		}
	}
	return 0;
}

LRESULT DateTimePicker::OnFormatNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMDATETIMEFORMAT pDetails = reinterpret_cast<LPNMDATETIMEFORMAT>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == DTN_FORMATA) {
			ATLASSERT_POINTER(pDetails, NMDATETIMEFORMATA);
			ATLASSERT_POINTER(pDetails->pszFormat, CHAR);
		} else {
			ATLASSERT_POINTER(pDetails, NMDATETIMEFORMATW);
			ATLASSERT_POINTER(pDetails->pszFormat, WCHAR);
		}
	#endif

	if(pDetails->pszFormat) {
		CComBSTR field;
		if(pNotificationDetails->code == DTN_FORMATA) {
			field = reinterpret_cast<LPNMDATETIMEFORMATA>(pDetails)->pszFormat;
		} else {
			field = reinterpret_cast<LPNMDATETIMEFORMATW>(pDetails)->pszFormat;
		}
		CComBSTR text;
		DATE date = 0;
		SystemTimeToVariantTime(&pDetails->st, &date);
		Raise_GetCallbackFieldText(field, date, &text);

		int requiredBufferSize = text.Length() + 1;
		if(requiredBufferSize > 64) {
			/* TODO: Set pszDisplay to a buffer that is large enough. The buffer must be valid until the next
			         call of DTN_FORMAT (for the same field?) or until the control is destroyed! */
			requiredBufferSize = 64;
		}
		if(pNotificationDetails->code == DTN_FORMATA) {
			lstrcpynA(const_cast<LPSTR>(reinterpret_cast<LPNMDATETIMEFORMATA>(pDetails)->pszDisplay), CW2A(text), requiredBufferSize);
		} else {
			lstrcpynW(const_cast<LPWSTR>(reinterpret_cast<LPNMDATETIMEFORMATW>(pDetails)->pszDisplay), text, requiredBufferSize);
		}
	}
	return 0;
}

LRESULT DateTimePicker::OnFormatQueryNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMDATETIMEFORMATQUERY pDetails = reinterpret_cast<LPNMDATETIMEFORMATQUERY>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == DTN_FORMATQUERYA) {
			ATLASSERT_POINTER(pDetails, NMDATETIMEFORMATQUERYA);
			ATLASSERT_POINTER(pDetails->pszFormat, CHAR);
		} else {
			ATLASSERT_POINTER(pDetails, NMDATETIMEFORMATQUERYW);
			ATLASSERT_POINTER(pDetails->pszFormat, WCHAR);
		}
	#endif

	if(pDetails->pszFormat) {
		CComBSTR field;
		if(pNotificationDetails->code == DTN_FORMATQUERYA) {
			field = reinterpret_cast<LPNMDATETIMEFORMATQUERYA>(pDetails)->pszFormat;
		} else {
			field = reinterpret_cast<LPNMDATETIMEFORMATQUERYW>(pDetails)->pszFormat;
		}
		Raise_GetCallbackFieldTextSize(field, &pDetails->szMax.cx, &pDetails->szMax.cy);
	}
	return 0;
}

LRESULT DateTimePicker::OnUserStringNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMDATETIMESTRING pDetails = reinterpret_cast<LPNMDATETIMESTRING>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == DTN_USERSTRINGA) {
			ATLASSERT_POINTER(pDetails, NMDATETIMESTRINGA);
			ATLASSERT_POINTER(pDetails->pszUserString, CHAR);
		} else {
			ATLASSERT_POINTER(pDetails, NMDATETIMESTRINGW);
			ATLASSERT_POINTER(pDetails->pszUserString, WCHAR);
		}
	#endif

	CComBSTR userInput;
	if(pNotificationDetails->code == DTN_USERSTRINGA) {
		userInput = reinterpret_cast<LPNMDATETIMESTRINGA>(pDetails)->pszUserString;
	} else {
		userInput = reinterpret_cast<LPNMDATETIMESTRINGW>(pDetails)->pszUserString;
	}
	DATE inputDate = 0;
	if(pDetails->dwFlags != GDT_NONE) {
		SystemTimeToVariantTime(&pDetails->st, &inputDate);
	}
	VARIANT_BOOL dateIsValid = BOOL2VARIANTBOOL(pDetails->dwFlags == GDT_VALID);

	Raise_ParseUserInput(userInput, &inputDate, &dateIsValid);

	if(dateIsValid != VARIANT_FALSE) {
		pDetails->dwFlags = (inputDate == 0 ? GDT_NONE : GDT_VALID);
		VariantTimeToSystemTime(inputDate, &pDetails->st);
	} else {
		pDetails->dwFlags = static_cast<DWORD>(GDT_ERROR);
	}
	return 0;
}

LRESULT DateTimePicker::OnWMKeyDownNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMDATETIMEWMKEYDOWN pDetails = reinterpret_cast<LPNMDATETIMEWMKEYDOWN>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == DTN_WMKEYDOWNA) {
			ATLASSERT_POINTER(pDetails, NMDATETIMEWMKEYDOWNA);
			ATLASSERT_POINTER(pDetails->pszFormat, CHAR);
		} else {
			ATLASSERT_POINTER(pDetails, NMDATETIMEWMKEYDOWNW);
			ATLASSERT_POINTER(pDetails->pszFormat, WCHAR);
		}
	#endif

	if(pDetails->pszFormat) {
		CComBSTR field;
		if(pNotificationDetails->code == DTN_WMKEYDOWNA) {
			field = reinterpret_cast<LPNMDATETIMEWMKEYDOWNA>(pDetails)->pszFormat;
		} else {
			field = reinterpret_cast<LPNMDATETIMEWMKEYDOWNW>(pDetails)->pszFormat;
		}
		SHORT keyCode = static_cast<SHORT>(pDetails->nVirtKey);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		DATE date = 0;
		SystemTimeToVariantTime(&pDetails->st, &date);
		Raise_CallbackFieldKeyDown(field, keyCode, shift, &date);
		VariantTimeToSystemTime(date, &pDetails->st);
	}
	return 0;
}

LRESULT DateTimePicker::OnCalendarGetDayStateNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMDAYSTATE pDetails = reinterpret_cast<LPNMDAYSTATE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMDAYSTATE);
	if(!pDetails) {
		return 0;
	}
	ATLASSERT(pDetails->cDayState > 0);

	int* pMonthDays = new int[pDetails->cDayState];
	pMonthDays[0] = GetMonthDays(&pDetails->stStart);
	ATLASSERT(pMonthDays[0] > 0);
	int totalNumberOfDates = pMonthDays[0] - (pDetails->stStart.wDay - 1);
	// add the other months
	SYSTEMTIME st = pDetails->stStart;
	for(int i = 1; i < pDetails->cDayState; ++i) {
		++st.wMonth;
		if(st.wMonth == 13) {
			st.wMonth = 1;
			++st.wYear;
		}
		pMonthDays[i] = GetMonthDays(&st);
		totalNumberOfDates += pMonthDays[i];
	}

	DATE firstDate;
	SystemTimeToVariantTime(&pDetails->stStart, &firstDate);
	LPSAFEARRAY pStates = SafeArrayCreateVectorEx(VT_BOOL, 0, totalNumberOfDates, NULL);
	ATLASSERT_POINTER(pStates, SAFEARRAY);
	VARIANT_BOOL b;
	LONG overallIndex = 0;
	/*for(int i = 0; i < pDetails->cDayState; ++i) {
		for(int day = 0; day < pMonthDays[i]; ++day, ++overallIndex) {
			b = BOOL2VARIANTBOOL(pDetails->prgDayState[i] & (0x00000001 << day));
			ATLVERIFY(SUCCEEDED(SafeArrayPutElement(pStates, &overallIndex, &b)));
		}
	}*/

	Raise_GetBoldDates(firstDate, totalNumberOfDates, &pStates);

	ZeroMemory(pDetails->prgDayState, pDetails->cDayState * sizeof(MONTHDAYSTATE));
	LONG l;
	ATLVERIFY(SUCCEEDED(SafeArrayGetLBound(pStates, 1, &l)));
	LONG u;
	ATLVERIFY(SUCCEEDED(SafeArrayGetUBound(pStates, 1, &u)));
	if(u >= l) {
		overallIndex = l;
		for(int i = 0; i < pDetails->cDayState; ++i) {
			for(int day = 0; day < pMonthDays[i] && overallIndex <= u; ++day, ++overallIndex) {
				ATLVERIFY(SUCCEEDED(SafeArrayGetElement(pStates, &overallIndex, &b)));
				if(b != VARIANT_FALSE) {
					pDetails->prgDayState[i] |= (0x00000001 << day);
				}
			}
		}
	}

	SafeArrayDestroy(pStates);
	delete[] pMonthDays;
	return 0;
}

LRESULT DateTimePicker::OnCalendarViewChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMVIEWCHANGE pDetails = reinterpret_cast<LPNMVIEWCHANGE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMVIEWCHANGE);
	if(!pDetails) {
		return 0;
	}

	Raise_CalendarViewChanged(static_cast<ViewConstants>(pDetails->dwOldView), static_cast<ViewConstants>(pDetails->dwNewView));
	return 0;
}


void DateTimePicker::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}

void DateTimePicker::ApplyCalendarFont(void)
{
	properties.calendarFont.dontGetFontObject = TRUE;
	//CWindow calendar = reinterpret_cast<HWND>(SendMessage(DTM_GETMONTHCAL, 0, 0));
	if(IsWindow()) {
		if(!properties.calendarFont.owningFontResource) {
			properties.calendarFont.currentFont.Detach();
		}
		properties.calendarFont.currentFont.Attach(NULL);

		if(properties.calendarUseSystemFont) {
			// use the system font
			NONCLIENTMETRICS nonClientMetrics = {0};
			nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
			properties.calendarFont.currentFont.CreateFontIndirect(&nonClientMetrics.lfMenuFont);
			properties.calendarFont.owningFontResource = TRUE;

			// apply the font
			SendMessage(DTM_SETMCFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.calendarFont.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.calendarFont.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.calendarFont.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.calendarFont.currentFont.Attach(hFont);
					properties.calendarFont.owningFontResource = FALSE;

					SendMessage(DTM_SETMCFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.calendarFont.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(DTM_SETMCFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(DTM_SETMCFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			//calendar.Invalidate();
		}
	}
	properties.calendarFont.dontGetFontObject = FALSE;
}


inline HRESULT DateTimePicker::Raise_CalendarClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		return Fire_CalendarClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarCloseUp(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CalendarCloseUp();
	//}
}

inline HRESULT DateTimePicker::Raise_CalendarContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		BOOL dontUsePosition = FALSE;
		DATE date = 0;
		UINT hitTestDetails = 0;
		if(x == -1 && y == -1) {
			// the event was caused by the keyboard
			dontUsePosition = TRUE;
			if(properties.processContextMenuKeys) {
				// TODO: If we find a way to retrieve the caret date's rectangle, then use it.
				// retrieve the caret item and propose its rectangle's middle as the menu's position
				WTL::CRect clientRectangle;
				containedCalendar.GetClientRect(&clientRectangle);
				WTL::CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
				hitTestDetails = htCalendar;

				SYSTEMTIME systemTime = {0};
				if(containedCalendar.SendMessage(MCM_GETCURSEL, 0, reinterpret_cast<LPARAM>(&systemTime))) {
					// SystemTimeToVariantTime doesn't like bogus values in these fields and MCM_GETCURSEL sometimes writes garbage to them
					systemTime.wMilliseconds = systemTime.wSecond = systemTime.wMinute = systemTime.wHour = 0;
					if(IsComctl32Version610OrNewer()) {
						switch(containedCalendar.SendMessage(MCM_GETCURRENTVIEW, 0, 0)) {
							case MCMV_CENTURY:
								systemTime.wYear = (systemTime.wYear / 10) * 10;
							case MCMV_DECADE:
								systemTime.wMonth = 1;
							case MCMV_YEAR:
								systemTime.wDay = 1;
								systemTime.wDayOfWeek = 0;
								break;
						}
					}
					SystemTimeToVariantTime(&systemTime, &date);
					hitTestDetails = htDate;

					if(IsComctl32Version610OrNewer()) {
						// try to get the date's rectangle - at first find the calendar index
						MCGRIDINFO gridInfo = {0};
						gridInfo.cbSize = sizeof(MCGRIDINFO);
						gridInfo.dwPart = MCGIP_CALENDAR;
						gridInfo.dwFlags = MCGIF_DATE;
						int calendarIndex = -1;
						for(gridInfo.iCalendar = 0; gridInfo.iCalendar < containedCalendar.SendMessage(MCM_GETCALENDARCOUNT, 0, 0); gridInfo.iCalendar++) {
							if(!containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
								break;
							}
							gridInfo.stStart.wMilliseconds = gridInfo.stStart.wSecond = gridInfo.stStart.wMinute = gridInfo.stStart.wHour = 0;
							DATE d = 0;
							SystemTimeToVariantTime(&gridInfo.stStart, &d);
							if(d > date) {
								calendarIndex = gridInfo.iCalendar - 1;
								break;
							} else {
								d = 0;
								SystemTimeToVariantTime(&gridInfo.stEnd, &d);
								if(date <= d) {
									calendarIndex = gridInfo.iCalendar;
									break;
								}
							}
						}
						if(calendarIndex >= 0) {
							gridInfo.iCalendar = calendarIndex;
							// now find the row and column index
							gridInfo.dwPart = MCGIP_CALENDARCELL;
							gridInfo.dwFlags = MCGIF_DATE;
							int columnIndex = -1;
							int rowIndex = -1;
							for(gridInfo.iCol = 0; gridInfo.iCol < 7; gridInfo.iCol++) {
								for(gridInfo.iRow = 0; gridInfo.iRow < 7; gridInfo.iRow++) {
									ZeroMemory(&gridInfo.stStart, sizeof(gridInfo.stStart));
									containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo));
									gridInfo.stStart.wMilliseconds = gridInfo.stStart.wSecond = gridInfo.stStart.wMinute = gridInfo.stStart.wHour = 0;
									DATE d = 0;
									SystemTimeToVariantTime(&gridInfo.stStart, &d);
									if(d == date) {
										columnIndex = gridInfo.iCol;
										rowIndex = gridInfo.iRow;
										break;
									}
								}
								if(columnIndex >= 0 && rowIndex >= 0) {
									break;
								}
							}
							if(columnIndex >= 0 && rowIndex >= 0) {
								gridInfo.dwFlags = MCGIF_RECT;
								gridInfo.iCol = columnIndex;
								gridInfo.iRow = rowIndex;
								if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
									WTL::CRect rc(&gridInfo.rc);
									WTL::CPoint centerPoint = rc.CenterPoint();
									x = centerPoint.x;
									y = centerPoint.y;
									dontUsePosition = FALSE;
									hitTestDetails = 0;
								}
							}
						}
					} else {
						// unfortunately it is not possible to retrieve the date's rectangle
					}
				}
			} else {
				return S_OK;
			}
		}

		if(!dontUsePosition) {
			date = CalendarHitTest(x, y, &hitTestDetails);
		}
		return Fire_CalendarContextMenu(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails), pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarDateMouseEnter(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus_Calendar.enteredControl) {
		return Fire_CalendarDateMouseEnter(date, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarDateMouseLeave(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus_Calendar.enteredControl) {
		return Fire_CalendarDateMouseLeave(date, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarDropDown(void)
{
	Raise_CreatedCalendarControlWindow(SendMessage(DTM_GETMONTHCAL, 0, 0));
	//if(m_nFreezeEvents == 0) {
		return Fire_CalendarDropDown();
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarKeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CalendarKeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarKeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CalendarKeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarKeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CalendarKeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		return Fire_CalendarMClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deCalendarMouseEvents)) {
		if(!mouseStatus_Calendar.enteredControl) {
			Raise_CalendarMouseEnter(button, shift, x, y);
		}
		if(!mouseStatus_Calendar.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = containedCalendar;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_CalendarMouseHover(button, shift, x, y);
		}
		mouseStatus_Calendar.StoreClickCandidate(button);
		if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
			containedCalendar.SetCapture();
		}

		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		return Fire_CalendarMouseDown(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	} else {
		if(!mouseStatus_Calendar.enteredControl) {
			mouseStatus_Calendar.EnterControl();
		}
		if(!mouseStatus_Calendar.hoveredControl) {
			mouseStatus_Calendar.HoverControl();
		}
		mouseStatus_Calendar.StoreClickCandidate(button);
		if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
			containedCalendar.SetCapture();
		}
		return S_OK;
	}
}

inline HRESULT DateTimePicker::Raise_CalendarMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = containedCalendar;
	trackingOptions.dwHoverTime = (properties.calendarHoverTime == -1 ? HOVER_DEFAULT : properties.calendarHoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus_Calendar.EnterControl();

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		dateUnderMouse = date;
		return Fire_CalendarMouseEnter(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return hr;
}

inline HRESULT DateTimePicker::Raise_CalendarMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus_Calendar.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deCalendarMouseEvents)) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		return Fire_CalendarMouseHover(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus_Calendar.enteredControl) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);

		if(dateUnderMouse) {
			Raise_CalendarDateMouseLeave(dateUnderMouse, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
			dateUnderMouse = 0;
		}
		mouseStatus_Calendar.LeaveControl();

		if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deCalendarMouseEvents)) {
			return Fire_CalendarMouseLeave(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
	}
	return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus_Calendar.enteredControl) {
		Raise_CalendarMouseEnter(button, shift, x, y);
	}
	mouseStatus_Calendar.lastPosition.x = x;
	mouseStatus_Calendar.lastPosition.y = y;

	UINT hitTestDetails = 0;
	DATE date = CalendarHitTest(x, y, &hitTestDetails);

	// Do we move over another date than before?
	if(date != dateUnderMouse) {
		if(dateUnderMouse) {
			Raise_CalendarDateMouseLeave(dateUnderMouse, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
		dateUnderMouse = date;
		if(date) {
			Raise_CalendarDateMouseEnter(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_CalendarMouseMove(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		return Fire_CalendarMouseWheel(date, button, shift, x, y, scrollAxis, wheelDelta, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deCalendarMouseEvents)) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		hr = Fire_CalendarMouseUp(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}

	if(mouseStatus_Calendar.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		containedCalendar.GetClientRect(&clientArea);
		containedCalendar.ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != containedCalendar) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(button != 1/*MouseButtonConstants.vbLeftButton*/ && GetCapture() == containedCalendar) {
			ReleaseCapture();
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deCalendarClickEvents)) {
						Raise_CalendarClick(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deCalendarClickEvents)) {
						Raise_CalendarRClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deCalendarClickEvents)) {
						Raise_CalendarMClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_CalendarXClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus_Calendar.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_CalendarOLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		containedCalendar.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

		UINT hitTestDetails = 0;
		DATE dropTarget = CalendarHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

		if(pData) {
			/* Actually we wouldn't need the next line, because the data object passed to this method should
				 always be the same as the data object that was passed to Raise_OLEDragEnter. */
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_CalendarOLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), dropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.CalendarOLEDragLeaveOrDrop();
	if(containedCalendar.IsWindow()) {
		containedCalendar.Invalidate();
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_CalendarOLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	containedCalendar.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.CalendarOLEDragEnter();

	UINT hitTestDetails = 0;
	DATE dropTarget = CalendarHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	LONG autoHScrollVelocity = 0;
	if(properties.calendarDragScrollTimeBase != 0) {
		BOOL isInScrollZone = ((hitTestDetails & MCHT_TITLEBTNNEXT) == MCHT_TITLEBTNNEXT || (hitTestDetails & MCHT_TITLEBTNPREV) == MCHT_TITLEBTNPREV);
		if(isInScrollZone) {
			// we're within the scroll zone - propose some velocity
			if((hitTestDetails & MCHT_TITLEBTNNEXT) == MCHT_TITLEBTNNEXT) {
				autoHScrollVelocity = 1;
			} else {
				autoHScrollVelocity = -1;
			}
		}
	}

	HRESULT hr = S_OK;
	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_CalendarOLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), dropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoHScrollVelocity);
		}
	//}

	if(properties.calendarDragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;

		LONG smallestInterval = abs(autoHScrollVelocity);
		if(smallestInterval) {
			smallestInterval = (properties.calendarDragScrollTimeBase == -1 ? GetDoubleClickTime() : properties.calendarDragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_CalendarOLEDragLeave(BOOL fakedLeave)
{
	KillTimer(timers.ID_DRAGSCROLL);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		UINT hitTestDetails = 0;
		DATE dropTarget = CalendarHitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);

		VARIANT_BOOL autoCloseUp = VARIANT_TRUE;
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_CalendarOLEDragLeave(dragDropStatus.pActiveDataObject, dropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoCloseUp);
		}
	//}

	if(fakedLeave) {
		dragDropStatus.isOverCalendar = FALSE;
	} else {
		dragDropStatus.pActiveDataObject = NULL;
		dragDropStatus.CalendarOLEDragLeaveOrDrop();
		containedCalendar.Invalidate();
	}
	if(autoCloseUp != VARIANT_FALSE) {
		CloseDropDownWindow();
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_CalendarOLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	containedCalendar.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	DATE dropTarget = CalendarHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	LONG autoHScrollVelocity = 0;
	if(properties.calendarDragScrollTimeBase != 0) {
		BOOL isInScrollZone = ((hitTestDetails & MCHT_TITLEBTNNEXT) == MCHT_TITLEBTNNEXT || (hitTestDetails & MCHT_TITLEBTNPREV) == MCHT_TITLEBTNPREV);
		if(isInScrollZone) {
			// we're within the scroll zone - propose some velocity
			if((hitTestDetails & MCHT_TITLEBTNNEXT) == MCHT_TITLEBTNNEXT) {
				autoHScrollVelocity = 1;
			} else {
				autoHScrollVelocity = -1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_CalendarOLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), dropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoHScrollVelocity);
		}
	//}

	if(properties.calendarDragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;

		LONG smallestInterval = abs(autoHScrollVelocity);
		if(smallestInterval) {
			smallestInterval = (properties.calendarDragScrollTimeBase == -1 ? GetDoubleClickTime() : properties.calendarDragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_CalendarRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		return Fire_CalendarRClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = CalendarHitTest(x, y, &hitTestDetails);
		return Fire_CalendarXClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CalendarViewChanged(ViewConstants oldView, ViewConstants newView)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CalendarViewChanged(oldView, newView);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CallbackFieldKeyDown(BSTR callbackField, SHORT keyCode, SHORT shift, DATE* pCurrentDate)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CallbackFieldKeyDown(callbackField, keyCode, shift, pCurrentDate);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// propose the middle of the control's client rectangle as the menu's position
				WTL::CRect clientRectangle;
				GetClientRect(&clientRectangle);
				WTL::CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
			} else {
				return S_OK;
			}
		}

		return Fire_ContextMenu(button, shift, x, y, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CreatedCalendarControlWindow(LONG hWndCalendar)
{
	// Wine does not create a new calendar window each time.
	if(containedCalendar != static_cast<HWND>(LongToHandle(hWndCalendar))) {
		// subclass the control
		containedCalendar.SubclassWindow(static_cast<HWND>(LongToHandle(hWndCalendar)));
	}

	containedCalendar.SendMessage(MCM_SETFIRSTDAYOFWEEK, 0, properties.startOfWeek);
	if(IsComctl32Version610OrNewer()) {
		containedCalendar.SendMessage(MCM_SETCURRENTVIEW, 0, properties.calendarView);
	} else {
		containedCalendar.ModifyStyle(MCS_MASK, properties.calendarStyle);
		WTL::CRect minimumRectangle;
		if(containedCalendar.SendMessage(MCM_GETMINREQRECT, 0, reinterpret_cast<LPARAM>(&minimumRectangle))) {
			// Wine moves the rectangle to (0,0), so don't assume we have a top-left position.
			//containedCalendar.MoveWindow(&minimumRectangle, FALSE);
			containedCalendar.SetWindowPos(NULL, &minimumRectangle, SWP_NOMOVE | SWP_NOOWNERZORDER | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOREDRAW);
		}
	}
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(containedCalendar, static_cast<IDropTarget*>(this)) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedCalendarControlWindow(hWndCalendar);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CreatedEditControlWindow(LONG hWndEdit)
{
	// subclass the control
	containedEdit.SubclassWindow(static_cast<HWND>(LongToHandle(hWndEdit)));

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedEditControlWindow(hWndEdit);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_CurrentDateChanged(VARIANT_BOOL dateSelected, DATE selectedDate)
{
	if(flags.noDateChangeEvent) {
		return S_OK;
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_CurrentDateChanged(dateSelected, selectedDate);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_DestroyedCalendarControlWindow(LONG hWndCalendar)
{
	if(!(properties.disabledEvents & deCalendarMouseEvents)) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = containedCalendar;
		trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		if(mouseStatus_Calendar.enteredControl) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			Raise_CalendarMouseLeave(button, shift, mouseStatus_Calendar.lastPosition.x, mouseStatus_Calendar.lastPosition.y);
		}
	}
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(containedCalendar) == S_OK);
	}

	//containedCalendar.UnsubclassWindow();     // done automatically
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedCalendarControlWindow(hWndCalendar);
	//}
}

inline HRESULT DateTimePicker::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	KillTimer(timers.ID_SETCUSTOMDATEFORMAT);
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_DestroyedEditControlWindow(LONG hWndEdit)
{
	//containedEdit.UnsubclassWindow();     // done automatically
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedEditControlWindow(hWndEdit);
	//}
}

inline HRESULT DateTimePicker::Raise_GetBoldDates(DATE firstDate, LONG numberOfDates, LPSAFEARRAY* ppStates)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetBoldDates(firstDate, numberOfDates, ppStates);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_GetCallbackFieldText(BSTR callbackField, DATE dateToFormat, BSTR* pTextToDisplay)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetCallbackFieldText(callbackField, dateToFormat, pTextToDisplay);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_GetCallbackFieldTextSize(BSTR callbackField, OLE_XSIZE_PIXELS* pTextWidth, OLE_YSIZE_PIXELS* pTextHeight)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetCallbackFieldTextSize(callbackField, pTextWidth, pTextHeight);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_MouseDown(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(index, button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			if(index == 2) {
				trackingOptions.hwndTrack = containedEdit;
			} else {
				trackingOptions.hwndTrack = *this;
			}
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		mouseStatus.StoreClickCandidate(button);
		if(index == 2) {
			containedEdit.SetCapture();
		} else {
			SetCapture();
		}

		return Fire_MouseDown(button, shift, x, y);
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
		mouseStatus.StoreClickCandidate(button);
		if(index == 2) {
			containedEdit.SetCapture();
		} else {
			SetCapture();
		}
		return S_OK;
	}
}

inline HRESULT DateTimePicker::Raise_MouseEnter(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	if(index == 2) {
		trackingOptions.hwndTrack = containedEdit;
	} else {
		trackingOptions.hwndTrack = *this;
	}
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(button, shift, x, y);
	//}
	//return hr;
}

inline HRESULT DateTimePicker::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT DateTimePicker::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus.enteredControl) {
		mouseStatus.LeaveControl();

		if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
			return Fire_MouseLeave(button, shift, x, y);
		}
	}
	return S_OK;
}

inline HRESULT DateTimePicker::Raise_MouseMove(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(index, button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_MouseUp(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y);
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		if(index == 2) {
			containedEdit.GetClientRect(&clientArea);
			containedEdit.ClientToScreen(&clientArea);
		} else {
			GetClientRect(&clientArea);
			ClientToScreen(&clientArea);
		}
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(index == 2) {
				if(WindowFromPoint(cursorPosition) != containedEdit) {
					hasLeftControl = TRUE;
				}
			} else {
				if(WindowFromPoint(cursorPosition) != *this) {
					hasLeftControl = TRUE;
				}
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(index == 2) {
			if(GetCapture() == containedEdit) {
				ReleaseCapture();
			}
		} else {
			if(GetCapture() == *this) {
				ReleaseCapture();
			}
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_MouseWheel(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(index, button, shift, x, y);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(button, shift, x, y, scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGDROPDOWN);
	dragDropStatus.dropDownTimerIsRunning = FALSE;

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

		if(pData) {
			/* Actually we wouldn't need the next line, because the data object passed to this method should
				 always be the same as the data object that was passed to Raise_OLEDragEnter. */
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT DateTimePicker::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	VARIANT_BOOL autoDropDown = VARIANT_FALSE;
	if(!(GetStyle() & DTS_UPDOWN) && !SendMessage(DTM_GETMONTHCAL, 0, 0)) {
		POINT pt = {mousePosition.x, mousePosition.y};
		if(IsComctl32Version610OrNewer()) {
			DATETIMEPICKERINFO controlInfo = {0};
			controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
			SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
			autoDropDown = BOOL2VARIANTBOOL(PtInRect(&controlInfo.rcButton, pt));
		} else {
			RECT buttonRectangle = {0};
			GetClientRect(&buttonRectangle);
			buttonRectangle.left = buttonRectangle.right - GetSystemMetrics(SM_CXVSCROLL);
			autoDropDown = BOOL2VARIANTBOOL(PtInRect(&buttonRectangle, pt));
		}
	}

	HRESULT hr = S_OK;
	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, &autoDropDown);
		}
	//}

	if(autoDropDown == VARIANT_FALSE) {
		// cancel automatic drop-down
		KillTimer(timers.ID_DRAGDROPDOWN);
		dragDropStatus.dropDownTimerIsRunning = FALSE;
	} else {
		if(properties.dragDropDownTime && !dragDropStatus.dropDownTimerIsRunning) {
			// start timered automatic drop-down
			dragDropStatus.dropDownTimerIsRunning = TRUE;
			SetTimer(timers.ID_DRAGDROPDOWN, (properties.dragDropDownTime == -1 ? GetDoubleClickTime() * 4 : properties.dragDropDownTime));
		}
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_OLEDragLeave(BOOL fakedLeave)
{
	KillTimer(timers.ID_DRAGDROPDOWN);
	dragDropStatus.dropDownTimerIsRunning = FALSE;

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y);
		}
	//}

	if(!fakedLeave) {
		dragDropStatus.pActiveDataObject = NULL;
		dragDropStatus.OLEDragLeaveOrDrop();
		Invalidate();
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	VARIANT_BOOL autoDropDown = VARIANT_FALSE;
	if(!(GetStyle() & DTS_UPDOWN) && !SendMessage(DTM_GETMONTHCAL, 0, 0)) {
		POINT pt = {mousePosition.x, mousePosition.y};
		if(IsComctl32Version610OrNewer()) {
			DATETIMEPICKERINFO controlInfo = {0};
			controlInfo.cbSize = sizeof(DATETIMEPICKERINFO);
			SendMessage(DTM_GETDATETIMEPICKERINFO, 0, reinterpret_cast<LPARAM>(&controlInfo));
			autoDropDown = BOOL2VARIANTBOOL(PtInRect(&controlInfo.rcButton, pt));
		} else {
			RECT buttonRectangle = {0};
			GetClientRect(&buttonRectangle);
			buttonRectangle.left = buttonRectangle.right - GetSystemMetrics(SM_CXVSCROLL);
			autoDropDown = BOOL2VARIANTBOOL(PtInRect(&buttonRectangle, pt));
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, &autoDropDown);
		}
	//}

	if(autoDropDown == VARIANT_FALSE) {
		// cancel automatic drop-down
		KillTimer(timers.ID_DRAGDROPDOWN);
		dragDropStatus.dropDownTimerIsRunning = FALSE;
	} else {
		if(properties.dragDropDownTime && !dragDropStatus.dropDownTimerIsRunning) {
			// start timered automatic drop-down
			dragDropStatus.dropDownTimerIsRunning = TRUE;
			SetTimer(timers.ID_DRAGDROPDOWN, (properties.dragDropDownTime == -1 ? GetDoubleClickTime() * 4 : properties.dragDropDownTime));
		}
	}

	return hr;
}

inline HRESULT DateTimePicker::Raise_ParseUserInput(BSTR userInput, DATE* pInputDate, VARIANT_BOOL* pDateIsValid)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ParseUserInput(userInput, pInputDate, pDateIsValid);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	} else {
		SetTimer(timers.ID_SETCUSTOMDATEFORMAT, timers.INT_SETCUSTOMDATEFORMAT);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT DateTimePicker::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y);
	//}
	//return S_OK;
}


void DateTimePicker::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

IMEModeConstants DateTimePicker::GetCurrentIMEContextMode(HWND hWnd)
{
	IMEModeConstants imeContextMode = imeNoControl;

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		imeContextMode = imeInherit;
	} else {
		HIMC hIMC = ImmGetContext(hWnd);
		if(hIMC) {
			if(ImmGetOpenStatus(hIMC)) {
				DWORD conversionMode = 0;
				DWORD sentenceMode = 0;
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				if(conversionMode & IME_CMODE_NATIVE) {
					if(conversionMode & IME_CMODE_KATAKANA) {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeKatakanaHalf];
						} else {
							imeContextMode = countryTable[imeAlphaFull];
						}
					} else {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeHiragana];
						} else {
							imeContextMode = countryTable[imeKatakana];
						}
					}
				} else {
					if(conversionMode & IME_CMODE_FULLSHAPE) {
						imeContextMode = countryTable[imeAlpha];
					} else {
						imeContextMode = countryTable[imeHangulFull];
					}
				}
			} else {
				imeContextMode = countryTable[imeDisable];
			}
			ImmReleaseContext(hWnd, hIMC);
		} else {
			imeContextMode = countryTable[imeOn];
		}
	}
	return imeContextMode;
}

void DateTimePicker::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
{
	if((IMEMode == imeInherit) || (IMEMode == imeNoControl) || !::IsWindow(hWnd)) {
		return;
	}

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		return;
	}

	// update IME mode
	HIMC hIMC = ImmGetContext(hWnd);
	if(IMEMode == imeDisable) {
		// disable IME
		if(hIMC) {
			// close the IME
			if(ImmGetOpenStatus(hIMC)) {
				ImmSetOpenStatus(hIMC, FALSE);
			}
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
		// remove the control's association to the IME context
		HIMC h = ImmAssociateContext(hWnd, NULL);
		if(h) {
			IMEFlags.hDefaultIMC = h;
		}
		return;
	} else {
		// enable IME
		if(!hIMC) {
			if(!IMEFlags.hDefaultIMC) {
				// create an IME context
				hIMC = ImmCreateContext();
				if(hIMC) {
					// associate the control with the IME context
					ImmAssociateContext(hWnd, hIMC);
				}
			} else {
				// associate the control with the default IME context
				ImmAssociateContext(hWnd, IMEFlags.hDefaultIMC);
			}
		} else {
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
	}

	hIMC = ImmGetContext(hWnd);
	if(hIMC) {
		DWORD conversionMode = 0;
		DWORD sentenceMode = 0;
		switch(IMEMode) {
			case imeOn:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				break;
			case imeOff:
				// close IME
				ImmSetOpenStatus(hIMC, FALSE);
				break;
			default:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				// switch conversion
				switch(IMEMode) {
					case imeHiragana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_KATAKANA;
						break;
					case imeKatakana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeKatakanaHalf:
						conversionMode |= (IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
					case imeAlphaFull:
						conversionMode |= IME_CMODE_FULLSHAPE;
						conversionMode &= ~(IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeAlpha:
						conversionMode |= IME_CMODE_ALPHANUMERIC;
						conversionMode &= ~(IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeHangulFull:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeHangul:
						conversionMode |= IME_CMODE_NATIVE;
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
				}
				ImmSetConversionStatus(hIMC, conversionMode, sentenceMode);
				break;
		}
		// each ImmGetContext() needs a ImmReleaseContext()
		ImmReleaseContext(hWnd, hIMC);
		hIMC = NULL;
	}
}

IMEModeConstants DateTimePicker::GetEffectiveIMEMode(void)
{
	IMEModeConstants IMEMode = properties.IMEMode;
	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

DWORD DateTimePicker::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD DateTimePicker::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.allowNullSelection) {
		style |= DTS_SHOWNONE;
	}
	switch(properties.style) {
		case sUpDown:
			style |= DTS_UPDOWN;
			break;
	}
	switch(properties.dateFormat) {
		case dfShortDate:
		case dfCustom:
			style |= DTS_SHORTDATEFORMAT;
			break;
		case dfShortDateWithFourDigitYear:
			style |= DTS_SHORTDATECENTURYFORMAT;
			break;
		case dfLongDate:
			style |= DTS_LONGDATEFORMAT;
			break;
		case dfTime:
			// DTS_TIMEFORMAT implies DTS_UPDOWN
			style |= DTS_TIMEFORMAT;
			properties.style = sUpDown;
			break;
	}
	if(!(properties.disabledEvents & deParseUserInput)) {
		style |= DTS_APPCANPARSE;
	}
	switch(properties.dropDownAlignment) {
		case ddaRight:
			style |= DTS_RIGHTALIGN;
			break;
	}
	return style;
}

void DateTimePicker::SendConfigurationMessages(void)
{
	/* HACK: SysDateTimePick32 will set WS_EX_CLIENTEDGE on creation, but we can remove it after the window
	         was created. */
	switch(properties.appearance) {
		case a2D:
			ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
		case a3D:
			ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
		case a3DLight:
			ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
	}

	DWORD calendarStyle = SendMessage(DTM_GETMCSTYLE, 0, 0);
	calendarStyle &= ~(MCS_NOTODAYCIRCLE | MCS_NOTODAY | MCS_WEEKNUMBERS | MCS_NOSELCHANGEONNAV | MCS_NOTRAILINGDATES | MCS_SHORTDAYSOFWEEK | MCS_DAYSTATE);
	if(!properties.calendarHighlightTodayDate) {
		calendarStyle |= MCS_NOTODAYCIRCLE;
	}
	if(!properties.calendarShowTodayDate) {
		calendarStyle |= MCS_NOTODAY;
	}
	if(properties.calendarShowWeekNumbers) {
		calendarStyle |= MCS_WEEKNUMBERS;
	}
	if(IsComctl32Version610OrNewer()) {
		if(properties.calendarKeepSelectionOnNavigation) {
			calendarStyle |= MCS_NOSELCHANGEONNAV;
		}
		if(!properties.calendarShowTrailingDates) {
			calendarStyle |= MCS_NOTRAILINGDATES;
		}
		if(properties.calendarUseShortestDayNames) {
			calendarStyle |= MCS_SHORTDAYSOFWEEK;
		}
	}
	//if(!(properties.disabledEvents & deGetBoldDates)) {
		calendarStyle |= MCS_DAYSTATE;
	//}
	SendMessage(DTM_SETMCSTYLE, 0, calendarStyle);

	SYSTEMTIME boundaryTimes[2];
	ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
	VariantTimeToSystemTime(properties.minDate, &boundaryTimes[0]);
	VariantTimeToSystemTime(properties.maxDate, &boundaryTimes[1]);
	DWORD flags = 0;
	if(!(boundaryTimes[0].wYear == 1752 && boundaryTimes[0].wMonth == 9 && boundaryTimes[0].wDay == 14)) {
		flags |= GDTR_MIN;
	}
	if(!(boundaryTimes[1].wYear == 9999 && boundaryTimes[1].wMonth == 12 && boundaryTimes[1].wDay == 31/* && boundaryTimes[1].wHour == 23 && boundaryTimes[1].wMinute == 59 && boundaryTimes[1].wSecond == 59*/)) {
		flags |= GDTR_MAX;
	}
	if(flags) {
		SendMessage(DTM_SETRANGE, flags, reinterpret_cast<LPARAM>(&boundaryTimes));
	}

	SendMessage(DTM_SETMCCOLOR, MCSC_BACKGROUND, OLECOLOR2COLORREF(properties.calendarBackColor));
	SendMessage(DTM_SETMCCOLOR, MCSC_TEXT, OLECOLOR2COLORREF(properties.calendarForeColor));
	SendMessage(DTM_SETMCCOLOR, MCSC_MONTHBK, OLECOLOR2COLORREF(properties.calendarMonthBackColor));
	SendMessage(DTM_SETMCCOLOR, MCSC_TITLEBK, OLECOLOR2COLORREF(properties.calendarTitleBackColor));
	SendMessage(DTM_SETMCCOLOR, MCSC_TITLETEXT, OLECOLOR2COLORREF(properties.calendarTitleForeColor));
	SendMessage(DTM_SETMCCOLOR, MCSC_TRAILINGTEXT, OLECOLOR2COLORREF(properties.calendarTrailingForeColor));

	if(properties.dateFormat == dfCustom) {
		COLE2T converter = properties.customDateFormat;
		LPTSTR pBuffer = converter;
		SendMessage(DTM_SETFORMAT, 0, reinterpret_cast<LPARAM>(pBuffer));
	}

	if(GetStyle() & DTS_SHOWNONE) {
		SYSTEMTIME systemTime = {0};
		if(SendMessage(DTM_GETSYSTEMTIME, 0, reinterpret_cast<LPARAM>(&systemTime)) != GDT_ERROR) {
			systemTime = properties.currentDate;
			if(SendMessage(DTM_SETSYSTEMTIME, (properties.dateSelected ? GDT_VALID : GDT_NONE), reinterpret_cast<LPARAM>(&systemTime))) {
				FireOnChanged(DISPID_DTP_CURRENTDATE);
				FireOnChanged(DISPID_DTP_DATESELECTED);
				FireOnChanged(DISPID_DTP_VALUE);
				SendOnDataChange();
			}
		}
	}

	ApplyFont();
}

HCURSOR DateTimePicker::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

DATE DateTimePicker::CalendarHitTest(LONG x, LONG y, PUINT pFlags)
{
	ATLASSERT(containedCalendar.IsWindow());

	MCHITTESTINFO hitTestInfo = {0};
	hitTestInfo.cbSize = RunTimeHelper::SizeOf_MCHITTESTINFO();
	hitTestInfo.pt.x = x;
	hitTestInfo.pt.y = y;
	UINT flags = containedCalendar.SendMessage(MCM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
	if(pFlags) {
		*pFlags = flags/*hitTestInfo.uHit*/;
	}
	DATE date = 0;
	if((flags & MCHT_CALENDAR) == MCHT_CALENDAR && (flags & MCHT_TODAYLINK) != MCHT_TODAYLINK) {
		SystemTimeToVariantTime(&hitTestInfo.st, &date);
	}
	return date;
}

DATE DateTimePicker::CalendarHitTest(LONG x, LONG y, PUINT pFlags, PINT pIndexOfHitCalendar, PINT pIndexOfHitColumn, PINT pIndexOfHitRow, LPRECT pCellRectangle)
{
	ATLASSERT(containedCalendar.IsWindow());

	MCHITTESTINFO hitTestInfo = {0};
	hitTestInfo.cbSize = RunTimeHelper::SizeOf_MCHITTESTINFO();
	hitTestInfo.pt.x = x;
	hitTestInfo.pt.y = y;
	UINT flags = containedCalendar.SendMessage(MCM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
	*pFlags = flags/*hitTestInfo.uHit*/;
	DATE date = 0;
	if((flags & MCHT_CALENDAR) == MCHT_CALENDAR && (flags & MCHT_TODAYLINK) != MCHT_TODAYLINK) {
		SystemTimeToVariantTime(&hitTestInfo.st, &date);
	}

	if(hitTestInfo.cbSize >= sizeof(MCHITTESTINFO)) {
		*pIndexOfHitCalendar = hitTestInfo.iOffset;
		*pIndexOfHitColumn = hitTestInfo.iCol;
		*pIndexOfHitRow = hitTestInfo.iRow;
		*pCellRectangle = hitTestInfo.rc;
	}

	return date;
}

BOOL DateTimePicker::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void DateTimePicker::AutoScroll(void)
{
	LONG realScrollTimeBase = properties.calendarDragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime();
	}

	if(dragDropStatus.autoScrolling.currentHScrollVelocity) {
		if(dragDropStatus.autoScrolling.currentHScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll to the left?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				// scroll left
				dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
				if(IsComctl32Version610OrNewer()) {
					MCGRIDINFO gridInfo = {0};
					gridInfo.cbSize = sizeof(MCGRIDINFO);
					gridInfo.dwFlags = MCGIF_RECT;
					gridInfo.dwPart = MCGIP_PREV;
					if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
						// this will take care of MCS_NOSELCHANGEONNAV
						containedCalendar.DefWindowProc(WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
						containedCalendar.DefWindowProc(WM_LBUTTONUP, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
					}
				} else {
					// MCS_NOSELCHANGEONNAV isn't supported anyway, so we don't need to calc the button rectangle
					containedCalendar.DefWindowProc(WM_KEYDOWN, VK_PRIOR, 0);
					containedCalendar.DefWindowProc(WM_KEYUP, VK_PRIOR, 0);
					// NOTE: Invalidate() doesn't work
					//containedCalendar.RedrawWindow();
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll to the right?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				// scroll right
				dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
				if(IsComctl32Version610OrNewer()) {
					MCGRIDINFO gridInfo = {0};
					gridInfo.cbSize = sizeof(MCGRIDINFO);
					gridInfo.dwFlags = MCGIF_RECT;
					gridInfo.dwPart = MCGIP_NEXT;
					if(containedCalendar.SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
						// this will take care of MCS_NOSELCHANGEONNAV
						containedCalendar.DefWindowProc(WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
						containedCalendar.DefWindowProc(WM_LBUTTONUP, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
					}
				} else {
					// MCS_NOSELCHANGEONNAV isn't supported anyway, so we don't need to calc the button rectangle
					containedCalendar.DefWindowProc(WM_KEYDOWN, VK_NEXT, 0);
					containedCalendar.DefWindowProc(WM_KEYUP, VK_NEXT, 0);
					// NOTE: Invalidate() doesn't work
					//containedCalendar.RedrawWindow();
				}
			}
		}
	}
}


BOOL DateTimePicker::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}