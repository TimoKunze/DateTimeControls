// Calendar.cpp: Superclasses SysMonthCal32.

#include "stdafx.h"
#include "Calendar.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC Calendar::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


Calendar::Calendar()
{
	properties.font.InitializePropertyWatcher(this, DISPID_CAL_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_CAL_MOUSEICON);

	SIZEL size = {173, 158};
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


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP Calendar::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ICalendar, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP Calendar::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 18 VT_I4 properties...
	pSize->QuadPart += 18 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 15 VT_BOOL properties...
	pSize->QuadPart += 15 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 2 VT_DATE properties...
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(DATE));

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
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

STDMETHODIMP Calendar::Load(LPSTREAM pStream)
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
		if(subSignature != 0x01010101/*4x 0x01 (-> Calendar)*/) {
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

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
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
	LONG dragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	VARTYPE vt;
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
	OLE_COLOR foreColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL highlightTodayDate = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL keepSelectionOnNavigation = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_DATE, &propertyValue))) {
		return hr;
	}
	DATE maxDate = propertyValue.date;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maxSelectedDates = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_DATE, &propertyValue))) {
		return hr;
	}
	DATE minDate = propertyValue.date;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR monthBackColor = propertyValue.lVal;

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
	VARIANT_BOOL multiSelect = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG padding = propertyValue.lVal;
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
	LONG scrollRate = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showTodayDate = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showTrailingDates = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showWeekNumbers = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	StartOfWeekConstants startOfWeek = static_cast<StartOfWeekConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR titleBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR titleForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR trailingForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL usePadding = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useShortestDayNames = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ViewConstants view = static_cast<ViewConstants>(propertyValue.lVal);

	VARIANT_BOOL detectDoubleClicks = VARIANT_FALSE;
	if(version >= 0x0102) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		detectDoubleClicks = propertyValue.boolVal;
	}


	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
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
	hr = put_DragScrollTimeBase(dragScrollTimeBase);
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
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HighlightTodayDate(highlightTodayDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_KeepSelectionOnNavigation(keepSelectionOnNavigation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxDate(maxDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaxSelectedDates(maxSelectedDates);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinDate(minDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MonthBackColor(monthBackColor);
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
	hr = put_MultiSelect(multiSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Padding(padding);
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
	hr = put_ScrollRate(scrollRate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowTodayDate(showTodayDate);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowTrailingDates(showTrailingDates);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowWeekNumbers(showWeekNumbers);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_StartOfWeek(startOfWeek);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TitleBackColor(titleBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TitleForeColor(titleForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TrailingForeColor(trailingForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UsePadding(usePadding);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseShortestDayNames(useShortestDayNames);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_View(view);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP Calendar::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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
	LONG subSignature = 0x01010101/*4x 0x01 (-> Calendar)*/;
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

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
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
	propertyValue.lVal = properties.dragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
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

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.highlightTodayDate);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.keepSelectionOnNavigation);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_DATE;
	propertyValue.date = properties.maxDate;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.maxSelectedDates;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_DATE;
	propertyValue.date = properties.minDate;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.monthBackColor;
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

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.padding;
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
	propertyValue.lVal = properties.scrollRate;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showTodayDate);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showTrailingDates);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showWeekNumbers);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.startOfWeek;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.titleBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.titleForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.trailingForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.usePadding);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useShortestDayNames);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.view;
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


HWND Calendar::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_DATE_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<Calendar>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT Calendar::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("Calendar ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void Calendar::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP Calendar::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<Calendar>::IOleObject_SetClientSite(pClientSite);

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

	return hr;
}

STDMETHODIMP Calendar::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL Calendar::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
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
	return CComControl<Calendar>::PreTranslateAccelerator(pMessage, hReturnValue);
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP Calendar::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP Calendar::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP Calendar::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
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

STDMETHODIMP Calendar::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
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
STDMETHODIMP Calendar::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
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

STDMETHODIMP Calendar::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_CAL_APPEARANCE:
		case DISPID_CAL_BORDERSTYLE:
		case DISPID_CAL_CALENDARGRIDCELLSELECTED:
		case DISPID_CAL_HIGHLIGHTTODAYDATE:
		case DISPID_CAL_MOUSEICON:
		case DISPID_CAL_MOUSEPOINTER:
		case DISPID_CAL_PADDING:
		case DISPID_CAL_SHOWTODAYDATE:
		case DISPID_CAL_SHOWTRAILINGDATES:
		case DISPID_CAL_SHOWWEEKNUMBERS:
		case DISPID_CAL_STARTOFWEEK:
		case DISPID_CAL_USEPADDING:
		case DISPID_CAL_USESHORTESTDAYNAMES:
		case DISPID_CAL_VIEW:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_CAL_DETECTDOUBLECLICKS:
		case DISPID_CAL_DISABLEDEVENTS:
		case DISPID_CAL_DONTREDRAW:
		case DISPID_CAL_HOVERTIME:
		case DISPID_CAL_KEEPSELECTIONONNAVIGATION:
		case DISPID_CAL_MAXDATE:
		case DISPID_CAL_MAXSELECTEDDATES:
		case DISPID_CAL_MINDATE:
		case DISPID_CAL_MULTISELECT:
		case DISPID_CAL_PROCESSCONTEXTMENUKEYS:
		case DISPID_CAL_RIGHTTOLEFT:
		case DISPID_CAL_SCROLLRATE:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_CAL_BACKCOLOR:
		case DISPID_CAL_FORECOLOR:
		case DISPID_CAL_MONTHBACKCOLOR:
		case DISPID_CAL_TITLEBACKCOLOR:
		case DISPID_CAL_TITLEFORECOLOR:
		case DISPID_CAL_TRAILINGFORECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_CAL_APPID:
		case DISPID_CAL_APPNAME:
		case DISPID_CAL_APPSHORTNAME:
		case DISPID_CAL_BUILD:
		case DISPID_CAL_CHARSET:
		case DISPID_CAL_HWND:
		case DISPID_CAL_ISRELEASE:
		case DISPID_CAL_PROGRAMMER:
		case DISPID_CAL_TESTER:
		case DISPID_CAL_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_CAL_DRAGSCROLLTIMEBASE:
		case DISPID_CAL_REGISTERFOROLEDRAGDROP:
		case DISPID_CAL_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_CAL_FONT:
		case DISPID_CAL_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_CAL_CALENDARCOUNT:
		case DISPID_CAL_CALENDARGRIDCELLTEXT:
		case DISPID_CAL_CARETDATE:
		case DISPID_CAL_CARETDATETEXT:
		case DISPID_CAL_ENABLED:
		case DISPID_CAL_FIRSTDATE:
		case DISPID_CAL_HEADERTEXT:
		case DISPID_CAL_LASTDATE:
		case DISPID_CAL_TODAYDATE:
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
CAtlString Calendar::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString Calendar::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString Calendar::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=DateTimeControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString Calendar::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString Calendar::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString Calendar::GetVersion(void)
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
STDMETHODIMP Calendar::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_CAL_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_CAL_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_CAL_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_CAL_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_CAL_STARTOFWEEK:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.startOfWeek), description);
			break;
		case DISPID_CAL_VIEW:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.view), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<Calendar>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP Calendar::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_CAL_BORDERSTYLE:
			c = 2;
			break;
		case DISPID_CAL_APPEARANCE:
			c = 3;
			break;
		case DISPID_CAL_RIGHTTOLEFT:
		case DISPID_CAL_VIEW:
			c = 4;
			break;
		case DISPID_CAL_STARTOFWEEK:
			c = 8;
			break;
		case DISPID_CAL_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<Calendar>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_CAL_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if(property == DISPID_CAL_STARTOFWEEK) {
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

STDMETHODIMP Calendar::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_CAL_APPEARANCE:
		case DISPID_CAL_BORDERSTYLE:
		case DISPID_CAL_MOUSEPOINTER:
		case DISPID_CAL_RIGHTTOLEFT:
		case DISPID_CAL_STARTOFWEEK:
		case DISPID_CAL_VIEW:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<Calendar>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP Calendar::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<Calendar>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT Calendar::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_CAL_APPEARANCE:
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
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_CAL_BORDERSTYLE:
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
		case DISPID_CAL_MOUSEPOINTER:
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
		case DISPID_CAL_RIGHTTOLEFT:
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
		case DISPID_CAL_STARTOFWEEK:
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
		case DISPID_CAL_VIEW:
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
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP Calendar::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 4;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StockColorPage;
		pPropertyPages->pElems[2] = CLSID_StockFontPage;
		pPropertyPages->pElems[3] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP Calendar::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<Calendar>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP Calendar::EnumVerbs(IEnumOLEVERB** ppEnumerator)
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
STDMETHODIMP Calendar::GetControlInfo(LPCONTROLINFO pControlInfo)
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

HRESULT Calendar::DoVerbAbout(HWND hWndParent)
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

HRESULT Calendar::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_CAL_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT Calendar::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP Calendar::get_Appearance(AppearanceConstants* pValue)
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

STDMETHODIMP Calendar::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_CAL_APPEARANCE);
	if(newValue < 0 || newValue >= aDefault) {
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
		FireOnChanged(DISPID_CAL_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 14;
	return S_OK;
}

STDMETHODIMP Calendar::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"DateTimeControls");
	return S_OK;
}

STDMETHODIMP Calendar::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"DTCtls");
	return S_OK;
}

STDMETHODIMP Calendar::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(MCM_GETCOLOR, MCSC_BACKGROUND, 0));
		if(color != OLECOLOR2COLORREF(properties.backColor)) {
			properties.backColor = color;
		}
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP Calendar::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CAL_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETCOLOR, MCSC_BACKGROUND, OLECOLOR2COLORREF(properties.backColor));
		}
		FireOnChanged(DISPID_CAL_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_BorderStyle(BorderStyleConstants* pValue)
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

STDMETHODIMP Calendar::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_CAL_BORDERSTYLE);
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
		FireOnChanged(DISPID_CAL_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP Calendar::get_CalendarCount(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		*pValue = SendMessage(MCM_GETCALENDARCOUNT, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::get_CalendarGridCellSelected(LONG columnIndex, LONG rowIndex, LONG calendarIndex/* = 0*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(calendarIndex < 0 || calendarIndex > 11) {
		return E_INVALIDARG;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDARCELL;
		gridInfo.dwFlags = MCGIF_DATE;
		gridInfo.iCalendar = calendarIndex;
		gridInfo.iCol = columnIndex;
		gridInfo.iRow = rowIndex;
		if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = BOOL2VARIANTBOOL(gridInfo.bSelected);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::get_CalendarGridCellText(LONG columnIndex, LONG rowIndex, LONG calendarIndex/* = 0*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}
	if(calendarIndex < 0 || calendarIndex > 11) {
		return E_INVALIDARG;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		WCHAR pBuffer[50];

		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDARCELL;
		gridInfo.dwFlags = MCGIF_NAME;
		gridInfo.iCalendar = calendarIndex;
		gridInfo.iCol = columnIndex;
		gridInfo.iRow = rowIndex;
		gridInfo.cchName = 50;
		gridInfo.pszName = pBuffer;
		if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = _bstr_t(gridInfo.pszName).Detach();
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::get_CaretDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	SYSTEMTIME systemTime = {0};
	if(IsWindow()) {
		if(SendMessage(MCM_GETCURSEL, 0, reinterpret_cast<LPARAM>(&systemTime))) {
			// SystemTimeToVariantTime doesn't like bogus values in these fields and MCM_GETCURSEL sometimes writes garbage to them
			systemTime.wMilliseconds = systemTime.wSecond = systemTime.wMinute = systemTime.wHour = 0;
			SystemTimeToVariantTime(&systemTime, pValue);
			return S_OK;
		}
	} else if(!(GetStyle() & MCS_MULTISELECT)) {
		GetSystemTime(&systemTime);
		systemTime.wMilliseconds = systemTime.wSecond = systemTime.wMinute = systemTime.wHour = 0;
		SystemTimeToVariantTime(&systemTime, pValue);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::put_CaretDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_CAL_CARETDATE);
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		SYSTEMTIME systemTime = {0};
		VariantTimeToSystemTime(newValue, &systemTime);
		if(SendMessage(MCM_SETCURSEL, 0, reinterpret_cast<LPARAM>(&systemTime))) {
			hr = S_OK;
			FireOnChanged(DISPID_CAL_CARETDATE);
		}
	}
	SendOnDataChange();
	return hr;
}

STDMETHODIMP Calendar::get_CaretDateText(LONG calendarIndex/* = 0*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}
	if(calendarIndex < 0 || calendarIndex > 11) {
		return E_INVALIDARG;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		WCHAR pBuffer[300];

		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDAR;
		gridInfo.dwFlags = MCGIF_NAME;
		gridInfo.iCalendar = calendarIndex;
		gridInfo.cchName = 300;
		gridInfo.pszName = pBuffer;
		if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = _bstr_t(gridInfo.pszName).Detach();
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::get_CharSet(BSTR* pValue)
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

STDMETHODIMP Calendar::get_DetectDoubleClicks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	return S_OK;
}

STDMETHODIMP Calendar::put_DetectDoubleClicks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_DETECTDOUBLECLICKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDoubleClicks != b) {
		properties.detectDoubleClicks = b;
		SetDirty(TRUE);

		FireOnChanged(DISPID_CAL_DETECTDOUBLECLICKS);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	/*if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & MCS_DAYSTATE) {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents & ~deGetBoldDates);
		} else {
			properties.disabledEvents = static_cast<DisabledEventsConstants>(properties.disabledEvents | deGetBoldDates);
		}
	}*/
	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP Calendar::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_CAL_DISABLEDEVENTS);
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
					dateUnderMouse = 0;
				}
			}
		}
		/*if((properties.disabledEvents & deGetBoldDates) != (newValue & deGetBoldDates)) {
			if(IsWindow()) {
				// MCS_DAYSTATE can't be set after control creation
				RecreateControlWindow();
			}
		}*/

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CAL_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP Calendar::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_CAL_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_DragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP Calendar::put_DragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CAL_DRAGSCROLLTIMEBASE);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragScrollTimeBase != newValue) {
		properties.dragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CAL_DRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_Enabled(VARIANT_BOOL* pValue)
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

STDMETHODIMP Calendar::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_ENABLED);
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

		FireOnChanged(DISPID_CAL_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_FirstDate(LONG rowIndex/* = -1*/, LONG calendarIndex/* = 0*/, DATE* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(calendarIndex < 0 || calendarIndex > 11) {
		return E_INVALIDARG;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = (rowIndex == -1 ? MCGIP_CALENDAR : MCGIP_CALENDARROW);
		gridInfo.dwFlags = MCGIF_DATE;
		gridInfo.iCalendar = calendarIndex;
		gridInfo.iRow = (rowIndex == -1 ? 0 : rowIndex);
		if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			SystemTimeToVariantTime(&gridInfo.stStart, pValue);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::get_Font(IFontDisp** ppFont)
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

STDMETHODIMP Calendar::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_CAL_FONT);
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
	FireOnChanged(DISPID_CAL_FONT);
	return S_OK;
}

STDMETHODIMP Calendar::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_CAL_FONT);
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
	FireOnChanged(DISPID_CAL_FONT);
	return S_OK;
}

STDMETHODIMP Calendar::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(MCM_GETCOLOR, MCSC_TEXT, 0));
		if(color != OLECOLOR2COLORREF(properties.foreColor)) {
			properties.foreColor = color;
		}
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP Calendar::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CAL_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETCOLOR, MCSC_TEXT, OLECOLOR2COLORREF(properties.foreColor));
		}
		FireOnChanged(DISPID_CAL_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_HeaderText(LONG calendarIndex/* = 0*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}
	if(calendarIndex < 0 || calendarIndex > 11) {
		return E_INVALIDARG;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		WCHAR pBuffer[300];

		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = MCGIP_CALENDARHEADER;
		gridInfo.dwFlags = MCGIF_NAME;
		gridInfo.iCalendar = calendarIndex;
		gridInfo.cchName = 300;
		gridInfo.pszName = pBuffer;
		if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			*pValue = _bstr_t(gridInfo.pszName).Detach();
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::get_HighlightTodayDate(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.highlightTodayDate = !(GetStyle() & MCS_NOTODAYCIRCLE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.highlightTodayDate);
	return S_OK;
}

STDMETHODIMP Calendar::put_HighlightTodayDate(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_HIGHLIGHTTODAYDATE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.highlightTodayDate != b) {
		properties.highlightTodayDate = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.highlightTodayDate) {
				ModifyStyle(MCS_NOTODAYCIRCLE, 0);
			} else {
				ModifyStyle(0, MCS_NOTODAYCIRCLE);
			}
		}
		FireOnChanged(DISPID_CAL_HIGHLIGHTTODAYDATE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP Calendar::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CAL_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CAL_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP Calendar::get_IsRelease(VARIANT_BOOL* pValue)
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

STDMETHODIMP Calendar::get_KeepSelectionOnNavigation(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.keepSelectionOnNavigation = ((GetStyle() & MCS_NOSELCHANGEONNAV) == MCS_NOSELCHANGEONNAV);
	}

	*pValue = BOOL2VARIANTBOOL(properties.keepSelectionOnNavigation);
	return S_OK;
}

STDMETHODIMP Calendar::put_KeepSelectionOnNavigation(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_KEEPSELECTIONONNAVIGATION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.keepSelectionOnNavigation != b) {
		properties.keepSelectionOnNavigation = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.keepSelectionOnNavigation) {
				ModifyStyle(0, MCS_NOSELCHANGEONNAV);
			} else {
				ModifyStyle(MCS_NOSELCHANGEONNAV, 0);
			}
		}
		FireOnChanged(DISPID_CAL_KEEPSELECTIONONNAVIGATION);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_LastDate(LONG rowIndex/* = -1*/, LONG calendarIndex/* = 0*/, DATE* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	if(calendarIndex < 0 || calendarIndex > 11) {
		return E_INVALIDARG;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = (rowIndex == -1 ? MCGIP_CALENDAR : MCGIP_CALENDARROW);
		gridInfo.dwFlags = MCGIF_DATE;
		gridInfo.iCalendar = calendarIndex;
		gridInfo.iRow = (rowIndex == -1 ? 0 : rowIndex);
		if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
			SystemTimeToVariantTime(&gridInfo.stEnd, pValue);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::get_MaxDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		DWORD setBoundaries = SendMessage(MCM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
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

STDMETHODIMP Calendar::put_MaxDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_CAL_MAXDATE);
	if(properties.maxDate != newValue) {
		properties.maxDate = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SYSTEMTIME boundaryTimes[2];
			ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
			// According to the documentation you can set min and max independently, but this does not seem to work.
			SendMessage(DTM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
			DWORD setBoundaries = SendMessage(MCM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
			if(!(setBoundaries & GDTR_MIN)) {
				boundaryTimes[0].wYear = 1752;
				boundaryTimes[0].wMonth = 9;
				boundaryTimes[0].wDay = 14;
			}
			VariantTimeToSystemTime(properties.maxDate, &boundaryTimes[1]);
			SendMessage(MCM_SETRANGE, GDTR_MIN | GDTR_MAX, reinterpret_cast<LPARAM>(&boundaryTimes));
		}
		FireOnChanged(DISPID_CAL_MAXDATE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_MaxSelectedDates(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.maxSelectedDates = static_cast<LONG>(SendMessage(MCM_GETMAXSELCOUNT, 0, 0));
	}

	*pValue = properties.maxSelectedDates;
	return S_OK;
}

STDMETHODIMP Calendar::put_MaxSelectedDates(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CAL_MAXSELECTEDDATES);
	if(newValue <= 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.maxSelectedDates != newValue) {
		properties.maxSelectedDates = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETMAXSELCOUNT, properties.maxSelectedDates, 0);
		}
		FireOnChanged(DISPID_CAL_MAXSELECTEDDATES);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_MinDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		DWORD setBoundaries = SendMessage(MCM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
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

STDMETHODIMP Calendar::put_MinDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_CAL_MINDATE);
	if(properties.minDate != newValue) {
		properties.minDate = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SYSTEMTIME boundaryTimes[2];
			ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
			// According to the documentation you can set min and max independently, but this does not seem to work.
			DWORD setBoundaries = SendMessage(MCM_GETRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes));
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
			SendMessage(MCM_SETRANGE, GDTR_MIN | GDTR_MAX, reinterpret_cast<LPARAM>(&boundaryTimes));
		}
		FireOnChanged(DISPID_CAL_MINDATE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_MonthBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(MCM_GETCOLOR, MCSC_MONTHBK, 0));
		if(color != OLECOLOR2COLORREF(properties.monthBackColor)) {
			properties.monthBackColor = color;
		}
	}

	*pValue = properties.monthBackColor;
	return S_OK;
}

STDMETHODIMP Calendar::put_MonthBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CAL_MONTHBACKCOLOR);
	if(properties.monthBackColor != newValue) {
		properties.monthBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETCOLOR, MCSC_MONTHBK, OLECOLOR2COLORREF(properties.monthBackColor));
		}
		FireOnChanged(DISPID_CAL_MONTHBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_MouseIcon(IPictureDisp** ppMouseIcon)
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

STDMETHODIMP Calendar::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_CAL_MOUSEICON);
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
	FireOnChanged(DISPID_CAL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP Calendar::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_CAL_MOUSEICON);
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
	FireOnChanged(DISPID_CAL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP Calendar::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP Calendar::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_CAL_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CAL_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_MultiSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiSelect = ((GetStyle() & MCS_MULTISELECT) == MCS_MULTISELECT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiSelect);
	return S_OK;
}

STDMETHODIMP Calendar::put_MultiSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_MULTISELECT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiSelect != b) {
		properties.multiSelect = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			/*if(properties.multiSelect) {
				ModifyStyle(0, MCS_MULTISELECT);
			} else {
				ModifyStyle(MCS_MULTISELECT, 0);
			}*/
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_CAL_MULTISELECT);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_Padding(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.padding = SendMessage(MCM_GETCALENDARBORDER, 0, 0);
	}

	*pValue = properties.padding;
	return S_OK;
}

STDMETHODIMP Calendar::put_Padding(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CAL_PADDING);
	if(properties.padding != newValue) {
		properties.padding = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			SendMessage(MCM_SETCALENDARBORDER, properties.usePadding, properties.padding);
		}
		FireOnChanged(DISPID_CAL_PADDING);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP Calendar::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CAL_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP Calendar::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP Calendar::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_REGISTERFOROLEDRAGDROP);
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
		}
		FireOnChanged(DISPID_CAL_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_RightToLeft(RightToLeftConstants* pValue)
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

STDMETHODIMP Calendar::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_CAL_RIGHTTOLEFT);
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
		FireOnChanged(DISPID_CAL_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_ScrollRate(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.scrollRate = static_cast<LONG>(SendMessage(MCM_GETMONTHDELTA, 0, 0));
	}

	*pValue = properties.scrollRate;
	return S_OK;
}

STDMETHODIMP Calendar::put_ScrollRate(LONG newValue)
{
	PUTPROPPROLOG(DISPID_CAL_SCROLLRATE);
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.scrollRate != newValue) {
		properties.scrollRate = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETMONTHDELTA, properties.scrollRate, 0);
		}
		FireOnChanged(DISPID_CAL_SCROLLRATE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_ShowTodayDate(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showTodayDate = !(GetStyle() & MCS_NOTODAY);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showTodayDate);
	return S_OK;
}

STDMETHODIMP Calendar::put_ShowTodayDate(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_SHOWTODAYDATE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showTodayDate != b) {
		properties.showTodayDate = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showTodayDate) {
				ModifyStyle(MCS_NOTODAY, 0);
			} else {
				ModifyStyle(0, MCS_NOTODAY);
			}
		}
		FireOnChanged(DISPID_CAL_SHOWTODAYDATE);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_ShowTrailingDates(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.showTrailingDates = !(GetStyle() & MCS_NOTRAILINGDATES);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showTrailingDates);
	return S_OK;
}

STDMETHODIMP Calendar::put_ShowTrailingDates(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_SHOWTRAILINGDATES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showTrailingDates != b) {
		properties.showTrailingDates = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.showTrailingDates) {
				ModifyStyle(MCS_NOTRAILINGDATES, 0);
			} else {
				ModifyStyle(0, MCS_NOTRAILINGDATES);
			}
		}
		FireOnChanged(DISPID_CAL_SHOWTRAILINGDATES);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_ShowWeekNumbers(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showWeekNumbers = ((GetStyle() & MCS_WEEKNUMBERS) == MCS_WEEKNUMBERS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showWeekNumbers);
	return S_OK;
}

STDMETHODIMP Calendar::put_ShowWeekNumbers(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_SHOWWEEKNUMBERS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showWeekNumbers != b) {
		properties.showWeekNumbers = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showWeekNumbers) {
				ModifyStyle(0, MCS_WEEKNUMBERS);
			} else {
				ModifyStyle(MCS_WEEKNUMBERS, 0);
			}
		}
		FireOnChanged(DISPID_CAL_SHOWWEEKNUMBERS);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_StartOfWeek(StartOfWeekConstants* pValue)
{
	ATLASSERT_POINTER(pValue, StartOfWeekConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD tmp = SendMessage(MCM_GETFIRSTDAYOFWEEK, 0, 0);
		if(HIWORD(tmp)) {
			properties.startOfWeek = static_cast<StartOfWeekConstants>(LOWORD(tmp));
		} else {
			properties.startOfWeek = sowUseLocale;
		}
	}

	*pValue = properties.startOfWeek;
	return S_OK;
}

STDMETHODIMP Calendar::put_StartOfWeek(StartOfWeekConstants newValue)
{
	PUTPROPPROLOG(DISPID_CAL_STARTOFWEEK);
	if(newValue < sowUseLocale || newValue > sowSunday) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.startOfWeek != newValue) {
		properties.startOfWeek = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETFIRSTDAYOFWEEK, 0, properties.startOfWeek);
		}
		FireOnChanged(DISPID_CAL_STARTOFWEEK);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP Calendar::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_CAL_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP Calendar::get_TitleBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(MCM_GETCOLOR, MCSC_TITLEBK, 0));
		if(color != OLECOLOR2COLORREF(properties.titleBackColor)) {
			properties.titleBackColor = color;
		}
	}

	*pValue = properties.titleBackColor;
	return S_OK;
}

STDMETHODIMP Calendar::put_TitleBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CAL_TITLEBACKCOLOR);
	if(properties.titleBackColor != newValue) {
		properties.titleBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETCOLOR, MCSC_TITLEBK, OLECOLOR2COLORREF(properties.titleBackColor));
		}
		FireOnChanged(DISPID_CAL_TITLEBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_TitleForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(MCM_GETCOLOR, MCSC_TITLETEXT, 0));
		if(color != OLECOLOR2COLORREF(properties.titleForeColor)) {
			properties.titleForeColor = color;
		}
	}

	*pValue = properties.titleForeColor;
	return S_OK;
}

STDMETHODIMP Calendar::put_TitleForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CAL_TITLEFORECOLOR);
	if(properties.titleForeColor != newValue) {
		properties.titleForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETCOLOR, MCSC_TITLETEXT, OLECOLOR2COLORREF(properties.titleForeColor));
		}
		FireOnChanged(DISPID_CAL_TITLEFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_TodayDate(DATE* pValue)
{
	ATLASSERT_POINTER(pValue, DATE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SYSTEMTIME systemTime = {0};
		if(SendMessage(MCM_GETTODAY, 0, reinterpret_cast<LPARAM>(&systemTime))) {
			systemTime.wMilliseconds = systemTime.wSecond = systemTime.wMinute = systemTime.wHour = 0;
			SystemTimeToVariantTime(&systemTime, pValue);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::put_TodayDate(DATE newValue)
{
	PUTPROPPROLOG(DISPID_CAL_TODAYDATE);
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		if(newValue == 0) {
			SendMessage(MCM_SETTODAY, 0, NULL);
		} else {
			SYSTEMTIME systemTime = {0};
			VariantTimeToSystemTime(newValue, &systemTime);
			SendMessage(MCM_SETTODAY, 0, reinterpret_cast<LPARAM>(&systemTime));
		}
		hr = S_OK;
		FireOnChanged(DISPID_CAL_TODAYDATE);
	}
	return hr;
}

STDMETHODIMP Calendar::get_TrailingForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(MCM_GETCOLOR, MCSC_TRAILINGTEXT, 0));
		if(color != OLECOLOR2COLORREF(properties.trailingForeColor)) {
			properties.trailingForeColor = color;
		}
	}

	*pValue = properties.trailingForeColor;
	return S_OK;
}

STDMETHODIMP Calendar::put_TrailingForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_CAL_TRAILINGFORECOLOR);
	if(properties.trailingForeColor != newValue) {
		properties.trailingForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(MCM_SETCOLOR, MCSC_TRAILINGTEXT, OLECOLOR2COLORREF(properties.trailingForeColor));
		}
		FireOnChanged(DISPID_CAL_TRAILINGFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_UsePadding(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.useShortestDayNames = ((GetStyle() & MCS_SHORTDAYSOFWEEK) == MCS_SHORTDAYSOFWEEK);
	}

	*pValue = BOOL2VARIANTBOOL(properties.usePadding);
	return S_OK;
}

STDMETHODIMP Calendar::put_UsePadding(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_USEPADDING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.usePadding != b) {
		properties.usePadding = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			//properties.padding = SendMessage(MCM_GETCALENDARBORDER, 0, 0);
			SendMessage(MCM_SETCALENDARBORDER, properties.usePadding, properties.padding);
		}
		FireOnChanged(DISPID_CAL_USEPADDING);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_UseShortestDayNames(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.useShortestDayNames = ((GetStyle() & MCS_SHORTDAYSOFWEEK) == MCS_SHORTDAYSOFWEEK);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useShortestDayNames);
	return S_OK;
}

STDMETHODIMP Calendar::put_UseShortestDayNames(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_USESHORTESTDAYNAMES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useShortestDayNames != b) {
		properties.useShortestDayNames = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.useShortestDayNames) {
				ModifyStyle(0, MCS_SHORTDAYSOFWEEK);
			} else {
				ModifyStyle(MCS_SHORTDAYSOFWEEK, 0);
			}
		}
		FireOnChanged(DISPID_CAL_USESHORTESTDAYNAMES);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP Calendar::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_CAL_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_CAL_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP Calendar::get_Version(BSTR* pValue)
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

STDMETHODIMP Calendar::get_View(ViewConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ViewConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(IsComctl32Version610OrNewer()) {
			properties.view = static_cast<ViewConstants>(SendMessage(MCM_GETCURRENTVIEW, 0, 0));
		} else {
			properties.view = vMonth;
		}
	}

	*pValue = properties.view;
	return S_OK;
}

STDMETHODIMP Calendar::put_View(ViewConstants newValue)
{
	PUTPROPPROLOG(DISPID_CAL_VIEW);
	if(newValue < MCMV_MONTH || newValue > MCMV_MAX) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.view != newValue) {
		properties.view = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			SendMessage(MCM_SETCURRENTVIEW, 0, properties.view);
		}
		FireOnChanged(DISPID_CAL_VIEW);
	}
	return S_OK;
}

STDMETHODIMP Calendar::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP Calendar::CountVisibleMonths(VARIANT_BOOL includePartialVisibleMonths/* = VARIANT_TRUE*/, DATE* pFirstVisibleDate/* = NULL*/, DATE* pLastVisibleDate/* = NULL*/, LONG* pVisibleMonths/* = NULL*/)
{
	ATLASSERT_POINTER(pVisibleMonths, LONG);
	if(!pVisibleMonths) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		*pVisibleMonths = SendMessage(MCM_GETMONTHRANGE, (includePartialVisibleMonths == VARIANT_FALSE ? GMR_VISIBLE : GMR_DAYSTATE), reinterpret_cast<LPARAM>(&boundaryTimes));
		if(pFirstVisibleDate) {
			SystemTimeToVariantTime(&boundaryTimes[0], pFirstVisibleDate);
		}
		if(pLastVisibleDate) {
			SystemTimeToVariantTime(&boundaryTimes[1], pLastVisibleDate);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::FinishOLEDragDrop(void)
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

STDMETHODIMP Calendar::GetMaxTodayWidth(OLE_XSIZE_PIXELS* pWidth)
{
	ATLASSERT_POINTER(pWidth, OLE_XSIZE_PIXELS);
	if(!pWidth) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pWidth = SendMessage(MCM_GETMAXTODAYWIDTH);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::GetMinimumSize(OLE_XSIZE_PIXELS* pWidth/* = NULL*/, OLE_YSIZE_PIXELS* pHeight/* = NULL*/)
{
	if(IsWindow()) {
		CRect minimumRectangle;
		if(SendMessage(MCM_GETMINREQRECT, 0, reinterpret_cast<LPARAM>(&minimumRectangle))) {
			if(pWidth) {
				*pWidth = minimumRectangle.Width();
				if(!(GetStyle() & MCS_NOTODAY)) {
					/* The rectangle returned by MCM_GETMINREQRECT does not include the width of the "Today" string,
					   if it is present. */
					int todayWidth = SendMessage(MCM_GETMAXTODAYWIDTH);
					*pWidth = max(*pWidth, todayWidth);
				}
			}
			if(pHeight) {
				*pHeight = minimumRectangle.Height();
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::GetRectangle(ControlPartConstants controlPart, LONG columnIndex/* = 0*/, LONG rowIndex/* = 0*/, LONG calendarIndex/* = 0*/, OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(IsWindow() && IsComctl32Version610OrNewer()) {
		MCGRIDINFO gridInfo = {0};
		gridInfo.cbSize = sizeof(MCGRIDINFO);
		gridInfo.dwPart = controlPart;
		gridInfo.dwFlags = MCGIF_RECT;
		gridInfo.iCalendar = calendarIndex;
		gridInfo.iCol = columnIndex;
		gridInfo.iRow = rowIndex;
		if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
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

STDMETHODIMP Calendar::GetSelection(DATE* pSelectionStart/* = NULL*/, DATE* pSelectionEnd/* = NULL*/)
{
	if(IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		if(SendMessage(MCM_GETSELRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes))) {
			if(pSelectionStart) {
				SystemTimeToVariantTime(&boundaryTimes[0], pSelectionStart);
			}
			if(pSelectionEnd) {
				SystemTimeToVariantTime(&boundaryTimes[1], pSelectionEnd);
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails/* = NULL*/, LONG* pIndexOfHitCalendar/* = NULL*/, LONG* pIndexOfHitColumn/* = NULL*/, LONG* pIndexOfHitRow/* = NULL*/, OLE_XPOS_PIXELS* pCellRectangleLeft/* = NULL*/, OLE_YPOS_PIXELS* pCellRectangleTop/* = NULL*/, OLE_XPOS_PIXELS* pCellRectangleRight/* = NULL*/, OLE_YPOS_PIXELS* pCellRectangleBottom/* = NULL*/, DATE* pHitDate/* = NULL*/)
{
	ATLASSERT_POINTER(pHitDate, DATE);
	if(!pHitDate) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT hitTestFlags = static_cast<UINT>(*pHitTestDetails);
		int hitCalendar = 0;
		int hitColumn = 0;
		int hitRow = 0;
		RECT hitCell = {0};

		*pHitDate = HitTest(x, y, &hitTestFlags, &hitCalendar, &hitColumn, &hitRow, &hitCell);
		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(hitTestFlags);
		}
		if(pIndexOfHitCalendar) {
			*pIndexOfHitCalendar = hitCalendar;
		}
		if(pIndexOfHitColumn) {
			*pIndexOfHitColumn = hitColumn;
		}
		if(pIndexOfHitRow) {
			*pIndexOfHitRow = hitRow;
		}
		if(pCellRectangleLeft) {
			*pCellRectangleLeft = hitCell.left;
		}
		if(pCellRectangleTop) {
			*pCellRectangleTop = hitCell.top;
		}
		if(pCellRectangleRight) {
			*pCellRectangleRight = hitCell.right;
		}
		if(pCellRectangleBottom) {
			*pCellRectangleBottom = hitCell.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP Calendar::MinimizeRectangle(OLE_XPOS_PIXELS* pLeft, OLE_YPOS_PIXELS* pTop, OLE_XPOS_PIXELS* pRight, OLE_YPOS_PIXELS* pBottom)
{
	ATLASSERT_POINTER(pLeft, OLE_XPOS_PIXELS);
	ATLASSERT_POINTER(pTop, OLE_YPOS_PIXELS);
	ATLASSERT_POINTER(pRight, OLE_XPOS_PIXELS);
	ATLASSERT_POINTER(pBottom, OLE_YPOS_PIXELS);
	if(!pLeft || !pTop || !pRight || !pBottom) {
		return E_POINTER;
	}

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		RECT minimumRectangle = {0};
		minimumRectangle.left = *pLeft;
		minimumRectangle.top = *pTop;
		minimumRectangle.right = *pRight;
		minimumRectangle.bottom = *pBottom;
		SendMessage(MCM_SIZERECTTOMIN, 0, reinterpret_cast<LPARAM>(&minimumRectangle));
		*pLeft = minimumRectangle.left;
		*pTop = minimumRectangle.top;
		*pRight = minimumRectangle.right;
		*pBottom = minimumRectangle.bottom;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP Calendar::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP Calendar::SetBoldDates(LPSAFEARRAY* ppStates, VARIANT_BOOL* pSucceeded)
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

	if(IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		int numberOfMonths = SendMessage(MCM_GETMONTHRANGE, GMR_DAYSTATE, reinterpret_cast<LPARAM>(&boundaryTimes));
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

		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(MCM_SETDAYSTATE, numberOfMonths, reinterpret_cast<LPARAM>(pMonthStates)));

		HeapFree(GetProcessHeap(), 0, pMonthStates);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Calendar::SetSelection(DATE selectionStart, DATE selectionEnd)
{
	if(IsWindow()) {
		SYSTEMTIME boundaryTimes[2];
		VariantTimeToSystemTime(selectionStart, &boundaryTimes[0]);
		VariantTimeToSystemTime(selectionEnd, &boundaryTimes[1]);
		if(SendMessage(MCM_SETSELRANGE, 0, reinterpret_cast<LPARAM>(&boundaryTimes))) {
			return S_OK;
		}
	}
	return E_FAIL;
}


LRESULT Calendar::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT Calendar::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
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

LRESULT Calendar::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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

LRESULT Calendar::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT Calendar::OnGetDlgCode(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam) | DLGC_WANTARROWS;
}

LRESULT Calendar::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT Calendar::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT Calendar::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<Calendar>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT Calendar::OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_LBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT Calendar::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_MBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT Calendar::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT Calendar::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT Calendar::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT Calendar::OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT Calendar::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT Calendar::OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_RBUTTONDOWN, wParam, lParam);
	}

	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT Calendar::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
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

LRESULT Calendar::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<Calendar>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<Calendar>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT Calendar::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_CAL_FONT) == S_FALSE) {
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
		FireOnChanged(DISPID_CAL_FONT);
	}

	return lr;
}

LRESULT Calendar::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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

LRESULT Calendar::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT Calendar::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
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

LRESULT Calendar::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
				ResetDayStates();
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		case timers.ID_INITDAYSTATES:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_INITDAYSTATES);
				ResetDayStates();
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		case timers.ID_DRAGSCROLL:
			AutoScroll();
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT Calendar::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;
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

LRESULT Calendar::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(!properties.detectDoubleClicks) {
		return SendMessage(WM_XBUTTONDOWN, wParam, lParam);
	}

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
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT Calendar::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Calendar::OnReflectedNotifyFormat(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
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

LRESULT Calendar::OnSetCalendarBorder(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	properties.usePadding = wParam;
	wasHandled = FALSE;
	return 0;
}

LRESULT Calendar::OnSetCurSel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_CAL_CARETDATE) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	FireOnChanged(DISPID_CAL_CARETDATE);
	SendOnDataChange();
	return lr;
}


LRESULT Calendar::OnGetDayStateNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT Calendar::OnSelChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMSELCHANGE pDetails = reinterpret_cast<LPNMSELCHANGE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMSELCHANGE);
	if(!pDetails) {
		return 0;
	}

	FireOnChanged(DISPID_CAL_CARETDATE);
	SendOnDataChange();

	DATE firstDate;
	DATE lastDate;
	SystemTimeToVariantTime(&pDetails->stSelStart, &firstDate);
	if(pDetails->stSelEnd.wYear) {
		SystemTimeToVariantTime(&pDetails->stSelEnd, &lastDate);
	} else {
		lastDate = firstDate;
	}
	Raise_SelectionChanged(firstDate, lastDate);
	return 0;
}

LRESULT Calendar::OnViewChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMVIEWCHANGE pDetails = reinterpret_cast<LPNMVIEWCHANGE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMVIEWCHANGE);
	if(!pDetails) {
		return 0;
	}

	Raise_ViewChanged(static_cast<ViewConstants>(pDetails->dwOldView), static_cast<ViewConstants>(pDetails->dwNewView));
	return 0;
}


void Calendar::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			NONCLIENTMETRICS nonClientMetrics = {0};
			nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
			properties.font.currentFont.CreateFontIndirect(&nonClientMetrics.lfMenuFont);
			properties.font.owningFontResource = TRUE;

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


inline HRESULT Calendar::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_Click(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		BOOL dontUsePosition = FALSE;
		DATE date = 0;
		UINT hitTestDetails = 0;
		if(x == -1 && y == -1) {
			// the event was caused by the keyboard
			dontUsePosition = TRUE;
			if(properties.processContextMenuKeys) {
				// retrieve the caret item and propose its rectangle's middle as the menu's position
				CRect clientRectangle;
				GetClientRect(&clientRectangle);
				CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
				hitTestDetails = htCalendar;

				SYSTEMTIME systemTime = {0};
				if(SendMessage(MCM_GETCURSEL, 0, reinterpret_cast<LPARAM>(&systemTime))) {
					// SystemTimeToVariantTime doesn't like bogus values in these fields and MCM_GETCURSEL sometimes writes garbage to them
					systemTime.wMilliseconds = systemTime.wSecond = systemTime.wMinute = systemTime.wHour = 0;
					if(IsComctl32Version610OrNewer()) {
						switch(SendMessage(MCM_GETCURRENTVIEW, 0, 0)) {
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
						for(gridInfo.iCalendar = 0; gridInfo.iCalendar < SendMessage(MCM_GETCALENDARCOUNT, 0, 0); gridInfo.iCalendar++) {
							if(!SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
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
									SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo));
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
								if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
									CRect rc(&gridInfo.rc);
									centerPoint = rc.CenterPoint();
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
			date = HitTest(x, y, &hitTestDetails);
		}
		return Fire_ContextMenu(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails), pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_DateMouseEnter(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_DateMouseEnter(date, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT Calendar::Raise_DateMouseLeave(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_DateMouseLeave(date, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT Calendar::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_DblClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	KillTimer(timers.ID_INITDAYSTATES);
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_GetBoldDates(DATE firstDate, LONG numberOfDates, LPSAFEARRAY* ppStates)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetBoldDates(firstDate, numberOfDates, ppStates);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_MClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_MDblClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	/* HACK: Fixes the Validate event of VB6. SysMonthCal32 doesn't seem to get the focus on WM_MOUSEACTIVATE
	 * by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired for some reason.
	 * Therefore we call SetFocus here.
	 */
	SetFocus();
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		mouseStatus.StoreClickCandidate(button);
		if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
			SetCapture();
		}

		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_MouseDown(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
		mouseStatus.StoreClickCandidate(button);
		if(button != 1/*MouseButtonConstants.vbLeftButton*/) {
			SetCapture();
		}
		return S_OK;
	}
}

inline HRESULT Calendar::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		dateUnderMouse = date;
		return Fire_MouseEnter(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return hr;
}

inline HRESULT Calendar::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_MouseHover(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT Calendar::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus.enteredControl) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);

		if(dateUnderMouse) {
			Raise_DateMouseLeave(dateUnderMouse, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
			dateUnderMouse = 0;
		}
		mouseStatus.LeaveControl();

		if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
			return Fire_MouseLeave(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
	}
	return S_OK;
}

inline HRESULT Calendar::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT hitTestDetails = 0;
	DATE date = HitTest(x, y, &hitTestDetails);

	// Do we move over another date than before?
	if(date != dateUnderMouse) {
		if(dateUnderMouse) {
			Raise_DateMouseLeave(dateUnderMouse, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
		dateUnderMouse = date;
		if(date) {
			Raise_DateMouseEnter(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		hr = Fire_MouseUp(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(button != 1/*MouseButtonConstants.vbLeftButton*/ && GetCapture() == *this) {
			ReleaseCapture();
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

inline HRESULT Calendar::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}

	UINT hitTestDetails = 0;
	DATE date = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(date, button, shift, x, y, scrollAxis, wheelDelta, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

		UINT hitTestDetails = 0;
		DATE dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

		if(pData) {
			/* Actually we wouldn't need the next line, because the data object passed to this method should
				 always be the same as the data object that was passed to Raise_OLEDragEnter. */
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), dropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT Calendar::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	DATE dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	LONG autoHScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
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
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), dropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoHScrollVelocity);
		}
	//}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;

		LONG smallestInterval = abs(autoHScrollVelocity);
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() : properties.dragScrollTimeBase) / smallestInterval;
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

inline HRESULT Calendar::Raise_OLEDragLeave(void)
{
	KillTimer(timers.ID_DRAGSCROLL);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		UINT hitTestDetails = 0;
		DATE dropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);

		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, dropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT Calendar::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	DATE dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	LONG autoHScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
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
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), dropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoHScrollVelocity);
		}
	//}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;

		LONG smallestInterval = abs(autoHScrollVelocity);
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() : properties.dragScrollTimeBase) / smallestInterval;
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

inline HRESULT Calendar::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_RClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_RDblClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	} else {
		SetTimer(timers.ID_INITDAYSTATES, timers.INT_INITDAYSTATES);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_SelectionChanged(DATE firstSelectedDate, DATE lastSelectedDate)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectionChanged(firstSelectedDate, lastSelectedDate);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_ViewChanged(ViewConstants oldView, ViewConstants newView)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ViewChanged(oldView, newView);
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_XClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT Calendar::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		DATE date = HitTest(x, y, &hitTestDetails);
		return Fire_XDblClick(date, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}


void Calendar::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

DWORD Calendar::GetExStyleBits(void)
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

DWORD Calendar::GetStyleBits(void)
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

	//if(!(properties.disabledEvents & deGetBoldDates)) {
		style |= MCS_DAYSTATE;
	//}
	if(!properties.highlightTodayDate) {
		style |= MCS_NOTODAYCIRCLE;
	}
	if(properties.multiSelect) {
		style |= MCS_MULTISELECT;
	}
	if(!properties.showTodayDate) {
		style |= MCS_NOTODAY;
	}
	if(properties.showWeekNumbers) {
		style |= MCS_WEEKNUMBERS;
	}
	if(IsComctl32Version610OrNewer()) {
		if(properties.keepSelectionOnNavigation) {
			style |= MCS_NOSELCHANGEONNAV;
		}
		if(!properties.showTrailingDates) {
			style |= MCS_NOTRAILINGDATES;
		}
		if(properties.useShortestDayNames) {
			style |= MCS_SHORTDAYSOFWEEK;
		}
	}
	return style;
}

void Calendar::SendConfigurationMessages(void)
{
	SendMessage(MCM_SETCURRENTVIEW, 0, properties.view);
	SendMessage(MCM_SETFIRSTDAYOFWEEK, 0, properties.startOfWeek);

	SYSTEMTIME boundaryTimes[2];
	ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
	VariantTimeToSystemTime(properties.minDate, &boundaryTimes[0]);
	VariantTimeToSystemTime(properties.maxDate, &boundaryTimes[1]);
	DWORD rangeFlags = 0;
	if(!(boundaryTimes[0].wYear == 1752 && boundaryTimes[0].wMonth == 9 && boundaryTimes[0].wDay == 14)) {
		rangeFlags |= GDTR_MIN;
	}
	if(!(boundaryTimes[1].wYear == 9999 && boundaryTimes[1].wMonth == 12 && boundaryTimes[1].wDay == 31/* && boundaryTimes[1].wHour == 23 && boundaryTimes[1].wMinute == 59 && boundaryTimes[1].wSecond == 59*/)) {
		rangeFlags |= GDTR_MAX;
	}
	if(rangeFlags) {
		SendMessage(MCM_SETRANGE, rangeFlags, reinterpret_cast<LPARAM>(&boundaryTimes));
	}

	SendMessage(MCM_SETCOLOR, MCSC_BACKGROUND, OLECOLOR2COLORREF(properties.backColor));
	SendMessage(MCM_SETCOLOR, MCSC_TEXT, OLECOLOR2COLORREF(properties.foreColor));
	SendMessage(MCM_SETCOLOR, MCSC_MONTHBK, OLECOLOR2COLORREF(properties.monthBackColor));
	SendMessage(MCM_SETCOLOR, MCSC_TITLEBK, OLECOLOR2COLORREF(properties.titleBackColor));
	SendMessage(MCM_SETCOLOR, MCSC_TITLETEXT, OLECOLOR2COLORREF(properties.titleForeColor));
	SendMessage(MCM_SETCOLOR, MCSC_TRAILINGTEXT, OLECOLOR2COLORREF(properties.trailingForeColor));
	SendMessage(MCM_SETMAXSELCOUNT, properties.maxSelectedDates, 0);
	SendMessage(MCM_SETMONTHDELTA, properties.scrollRate, 0);
	SendMessage(MCM_SETCALENDARBORDER, properties.usePadding, properties.padding);

	ApplyFont();
}

HCURSOR Calendar::MousePointerConst2hCursor(MousePointerConstants mousePointer)
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

DATE Calendar::HitTest(LONG x, LONG y, PUINT pFlags)
{
	ATLASSERT(IsWindow());

	MCHITTESTINFO hitTestInfo = {0};
	hitTestInfo.cbSize = RunTimeHelper::SizeOf_MCHITTESTINFO();
	hitTestInfo.pt.x = x;
	hitTestInfo.pt.y = y;
	UINT hitTestFlags = SendMessage(MCM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
	if(pFlags) {
		*pFlags = hitTestFlags/*hitTestInfo.uHit*/;
	}
	DATE date = 0;
	if((hitTestFlags & MCHT_CALENDAR) == MCHT_CALENDAR && (hitTestFlags & MCHT_TODAYLINK) != MCHT_TODAYLINK) {
		SystemTimeToVariantTime(&hitTestInfo.st, &date);
	}
	return date;
}

DATE Calendar::HitTest(LONG x, LONG y, PUINT pFlags, PINT pIndexOfHitCalendar, PINT pIndexOfHitColumn, PINT pIndexOfHitRow, LPRECT pCellRectangle)
{
	ATLASSERT(IsWindow());

	MCHITTESTINFO hitTestInfo = {0};
	hitTestInfo.cbSize = RunTimeHelper::SizeOf_MCHITTESTINFO();
	hitTestInfo.pt.x = x;
	hitTestInfo.pt.y = y;
	UINT hitTestFlags = SendMessage(MCM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
	*pFlags = hitTestFlags/*hitTestInfo.uHit*/;
	DATE date = 0;
	if((hitTestFlags & MCHT_CALENDAR) == MCHT_CALENDAR && (hitTestFlags & MCHT_TODAYLINK) != MCHT_TODAYLINK) {
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

BOOL Calendar::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void Calendar::AutoScroll(void)
{
	LONG realScrollTimeBase = properties.dragScrollTimeBase;
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
					if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
						// this will take care of MCS_NOSELCHANGEONNAV
						DefWindowProc(WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
						DefWindowProc(WM_LBUTTONUP, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
					}
				} else {
					// MCS_NOSELCHANGEONNAV isn't supported anyway, so we don't need to calc the button rectangle
					DefWindowProc(WM_KEYDOWN, VK_PRIOR, 0);
					DefWindowProc(WM_KEYUP, VK_PRIOR, 0);
					// NOTE: Invalidate() doesn't work
					//RedrawWindow();
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
					if(SendMessage(MCM_GETCALENDARGRIDINFO, 0, reinterpret_cast<LPARAM>(&gridInfo))) {
						// this will take care of MCS_NOSELCHANGEONNAV
						DefWindowProc(WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
						DefWindowProc(WM_LBUTTONUP, MK_LBUTTON, MAKELPARAM(gridInfo.rc.left, gridInfo.rc.top));
					}
				} else {
					// MCS_NOSELCHANGEONNAV isn't supported anyway, so we don't need to calc the button rectangle
					DefWindowProc(WM_KEYDOWN, VK_NEXT, 0);
					DefWindowProc(WM_KEYUP, VK_NEXT, 0);
					// NOTE: Invalidate() doesn't work
					//RedrawWindow();
				}
			}
		}
	}
}

void Calendar::ResetDayStates(void)
{
	if(GetStyle() & MCS_DAYSTATE) {
		/* HACK: For some reason the control starts with random bold states and doesn't send
		   MCN_GETDAYSTATE. */
		SYSTEMTIME boundaryTimes[2];
		ZeroMemory(boundaryTimes, 2 * sizeof(SYSTEMTIME));
		int numberOfMonths = SendMessage(MCM_GETMONTHRANGE, GMR_DAYSTATE, reinterpret_cast<LPARAM>(&boundaryTimes));
		MONTHDAYSTATE* pMonths = new MONTHDAYSTATE[numberOfMonths];

		NMDAYSTATE notificationDetails = {0};
		notificationDetails.nmhdr.code = MCN_GETDAYSTATE;
		notificationDetails.nmhdr.idFrom = GetDlgCtrlID();
		notificationDetails.nmhdr.hwndFrom = *this;
		notificationDetails.stStart = boundaryTimes[0];
		notificationDetails.stStart.wDay = 1;
		notificationDetails.cDayState = numberOfMonths;
		notificationDetails.prgDayState = pMonths;
		GetParent().SendMessage(WM_NOTIFY, notificationDetails.nmhdr.idFrom, reinterpret_cast<LPARAM>(&notificationDetails));

		SendMessage(MCM_SETDAYSTATE, numberOfMonths, reinterpret_cast<LPARAM>(pMonths));
		delete[] pMonths;
	}
}


BOOL Calendar::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}