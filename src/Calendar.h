//////////////////////////////////////////////////////////////////////
/// \class Calendar
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c SysMonthCal32</em>
///
/// This class superclasses \c SysMonthCal32 and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa DTCtlsLibU::ICalendar
/// \else
///   \sa DTCtlsLibA::ICalendar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "DTCtlsU.h"
#else
	#include "DTCtlsA.h"
#endif
#include "_ICalendarEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "CommonProperties.h"
#include "TargetOLEDataObject.h"


class ATL_NO_VTABLE Calendar : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<ICalendar, &IID_ICalendar, &LIBID_DTCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<ICalendar, &IID_ICalendar, &LIBID_DTCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<Calendar>,
    public IOleControlImpl<Calendar>,
    public IOleObjectImpl<Calendar>,
    public IOleInPlaceActiveObjectImpl<Calendar>,
    public IViewObjectExImpl<Calendar>,
    public IOleInPlaceObjectWindowlessImpl<Calendar>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<Calendar>,
    public Proxy_ICalendarEvents<Calendar>,
    public IPersistStorageImpl<Calendar>,
    public IPersistPropertyBagImpl<Calendar>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<Calendar>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_Calendar, &__uuidof(_ICalendarEvents), &LIBID_DTCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_Calendar, &__uuidof(_ICalendarEvents), &LIBID_DTCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<Calendar>,
    public CComCoClass<Calendar, &CLSID_Calendar>,
    public CComControl<Calendar>,
    public IPerPropertyBrowsingImpl<Calendar>,
    public IDropTarget,
    public ICategorizeProperties,
    public ICreditsProvider
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	Calendar();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_CALENDAR)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("CalendarU"), MONTHCAL_CLASSW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("CalendarA"), MONTHCAL_CLASSA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(Calendar)
			COM_INTERFACE_ENTRY(ICalendar)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
		END_COM_MAP()

		BEGIN_PROP_MAP(Calendar)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("Appearance", DISPID_CAL_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BackColor", DISPID_CAL_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_CAL_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DetectDoubleClicks", DISPID_CAL_DETECTDOUBLECLICKS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_CAL_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_CAL_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragScrollTimeBase", DISPID_CAL_DRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_CAL_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_CAL_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_CAL_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HighlightTodayDate", DISPID_CAL_HIGHLIGHTTODAYDATE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("HoverTime", DISPID_CAL_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("KeepSelectionOnNavigation", DISPID_CAL_KEEPSELECTIONONNAVIGATION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MaxDate", DISPID_CAL_MAXDATE, CLSID_NULL, VT_DATE)
			PROP_ENTRY_TYPE("MaxSelectedDates", DISPID_CAL_MAXSELECTEDDATES, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MinDate", DISPID_CAL_MINDATE, CLSID_NULL, VT_DATE)
			PROP_ENTRY_TYPE("MonthBackColor", DISPID_CAL_MONTHBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_CAL_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_CAL_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MultiSelect", DISPID_CAL_MULTISELECT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Padding", DISPID_CAL_PADDING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_CAL_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_CAL_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_CAL_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ScrollRate", DISPID_CAL_SCROLLRATE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ShowTodayDate", DISPID_CAL_SHOWTODAYDATE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ShowTrailingDates", DISPID_CAL_SHOWTRAILINGDATES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ShowWeekNumbers", DISPID_CAL_SHOWWEEKNUMBERS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("StartOfWeek", DISPID_CAL_STARTOFWEEK, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_CAL_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("TitleBackColor", DISPID_CAL_TITLEBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("TitleForeColor", DISPID_CAL_TITLEFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("TrailingForeColor", DISPID_CAL_TRAILINGFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("UsePadding", DISPID_CAL_USEPADDING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseShortestDayNames", DISPID_CAL_USESHORTESTDAYNAMES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_CAL_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("View", DISPID_CAL_VIEW, CLSID_NULL, VT_I4)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(Calendar)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_ICalendarEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(Calendar)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_GETDLGCODE, OnGetDlgCode)
			MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

			MESSAGE_HANDLER(MCM_SETCALENDARBORDER, OnSetCalendarBorder)
			MESSAGE_HANDLER(MCM_SETCURSEL, OnSetCurSel)

			REFLECTED_NOTIFY_CODE_HANDLER(MCN_GETDAYSTATE, OnGetDayStateNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(MCN_SELCHANGE, OnSelChangeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(MCN_VIEWCHANGE, OnViewChangeNotification)

			CHAIN_MSG_MAP(CComControl<Calendar>)
		END_MSG_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICalendar
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by
	/// the \c AppearanceConstants enumeration except \c aDefault is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Appearance, get_BorderStyle, get_Padding, DTCtlsLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, get_Padding, DTCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration except \c aDefault is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Appearance, put_BorderStyle, put_Padding, DTCtlsLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, put_Padding, DTCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BackColor, get_ForeColor, get_TitleBackColor, get_TitleForeColor, get_MonthBackColor,
	///     get_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BackColor, put_ForeColor, put_TitleBackColor, put_TitleForeColor, put_MonthBackColor,
	///     put_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values defined
	/// by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, get_Appearance, get_Padding, DTCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, get_Padding, DTCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values defined by
	/// the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, put_Appearance, put_Padding, DTCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, put_Padding, DTCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarCount property</em>
	///
	/// Retrieves the number of calendars currently displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa MinimizeRectangle
	virtual HRESULT STDMETHODCALLTYPE get_CalendarCount(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarGridCellSelected property</em>
	///
	/// Retrieves whether the specified calendar grid cell is selected. If set to \c VARIANT_TRUE, the cell
	/// is selected; otherwise not.
	///
	/// \param[in] columnIndex The column index of the grid cell for which to retrieve the selection state.
	///            The index starts at -1, which identifies the week numbers column.
	/// \param[in] rowIndex The row index of the grid cell for which to retrieve the selection state. The
	///            index starts at -1, which identifies the week days row.
	/// \param[in] calendarIndex The zero-based index of the calendar for which to retrieve the selection
	///            state. The control can display up to 12 calendars at once.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CaretDate, get_MultiSelect, GetSelection, SetSelection, get_CalendarGridCellText
	virtual HRESULT STDMETHODCALLTYPE get_CalendarGridCellSelected(LONG columnIndex, LONG rowIndex, LONG calendarIndex = 0, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c CalendarGridCellText property</em>
	///
	/// Retrieves the text being displayed in the specified calendar grid cell. The string contains the
	/// Unicode character 8206 before each part of the text, e. g. &ldquo;?January, ?2009&rdquo; (? stands
	/// for character 8206). This marker can be used to extract specific parts from the text.
	///
	/// \param[in] columnIndex The column index of the grid cell for which to retrieve the text. The index
	///            starts at -1, which identifies the week numbers column.
	/// \param[in] rowIndex The row index of the grid cell for which to retrieve the text. The index starts
	///            at -1, which identifies the week days row.
	/// \param[in] calendarIndex The zero-based index of the calendar for which to retrieve the text. The
	///            control can display up to 12 calendars at once.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_HeaderText, get_CaretDateText, get_CalendarGridCellSelected
	virtual HRESULT STDMETHODCALLTYPE get_CalendarGridCellText(LONG columnIndex, LONG rowIndex, LONG calendarIndex = 0, BSTR* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c CaretDate property</em>
	///
	/// Retrieves the currently selected date.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa put_CaretDate, get_MultiSelect, get_MinDate, get_MaxDate, get_CaretDateText,
	///     Raise_SelectionChanged
	virtual HRESULT STDMETHODCALLTYPE get_CaretDate(DATE* pValue);
	/// \brief <em>Sets the \c CaretDate property</em>
	///
	/// Sets the currently selected date.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa get_CaretDate, put_MultiSelect, put_MinDate, put_MaxDate, get_CaretDateText,
	///     Raise_SelectionChanged
	virtual HRESULT STDMETHODCALLTYPE put_CaretDate(DATE newValue);
	/// \brief <em>Retrieves the current setting of the \c CaretDateText property</em>
	///
	/// Retrieves the text representing the caret date. The string contains the Unicode character 8206
	/// before each part of the text, e. g. &ldquo;?January, ?2009&rdquo; (? stands for character 8206).
	/// This marker can be used to extract specific parts from the text.
	///
	/// \param[in] calendarIndex The zero-based index of the calendar for which to retrieve the text. The
	///            control can display up to 12 calendars at once.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CaretDate, get_CalendarGridCellText, get_HeaderText
	virtual HRESULT STDMETHODCALLTYPE get_CaretDateText(LONG calendarIndex = 0, BSTR* pValue = NULL);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c DetectDoubleClicks property</em>
	///
	/// Retrieves whether double clicks are enabled or disabled. If set to \c VARIANT_TRUE, double clicks are
	/// accepted; otherwise all clicks are handled as single clicks.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Enabling double-clicks may lead to accidental double-clicks, e.g. if the user rapidly
	///          clicks onto the navigation arrows for selecting another month.
	///
	/// \sa put_DetectDoubleClicks, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	virtual HRESULT STDMETHODCALLTYPE get_DetectDoubleClicks(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DetectDoubleClicks property</em>
	///
	/// Enables or disables double clicks. If set to \c VARIANT_TRUE, double clicks are accepted; otherwise
	/// all clicks are handled as single clicks.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Enabling double-clicks may lead to accidental double-clicks, e.g. if the user rapidly
	///          clicks onto the navigation arrows for selecting another month.
	///
	/// \sa get_DetectDoubleClicks, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	virtual HRESULT STDMETHODCALLTYPE put_DetectDoubleClicks(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, DTCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, DTCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, DTCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, DTCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DontRedraw property</em>
	///
	/// Retrieves whether automatic redrawing of the control is enabled or disabled. If set to
	/// \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE get_DontRedraw(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DontRedraw property</em>
	///
	/// Enables or disables automatic redrawing of the control. Disabling redraw while doing large changes
	/// on the control may increase performance. If set to \c VARIANT_FALSE, the control will redraw itself
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE put_DontRedraw(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragScrollTimeBase, get_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c DragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling during a drag'n'drop operation. If set to 0, auto-scrolling is
	/// disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragScrollTimeBase, put_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragScrollTimeBase(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FirstDate property</em>
	///
	/// Retrieves the specified calendar's first date that belongs to the month it is displaying.
	///
	/// \param[in] rowIndex The zero-based row index for which to retrieve the date. If set to -1, the first
	///            date of the whole calendar is retrieved.
	/// \param[in] calendarIndex The zero-based index of the calendar for which to retrieve the date. The
	///            control can display up to 12 calendars at once.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_LastDate, get_MinDate
	virtual HRESULT STDMETHODCALLTYPE get_FirstDate(LONG rowIndex = -1, LONG calendarIndex = 0, DATE* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the control's content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.\n
	///          Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_Font, put_Font, put_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_ForeColor, get_BackColor, get_TitleBackColor, get_TitleForeColor, get_MonthBackColor,
	///     get_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_ForeColor, put_BackColor, put_TitleBackColor, put_TitleForeColor, put_MonthBackColor,
	///     put_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c HeaderText property</em>
	///
	/// Retrieves the text being displayed in the calendar header, for instance &ldquo;January, 2009&rdquo;.
	/// The string contains the Unicode character 8206 before each part of the text, e. g. &ldquo;?January,
	/// ?2009&rdquo; (? stands for character 8206). This marker can be used to extract specific parts from
	/// the text.
	///
	/// \param[in] calendarIndex The zero-based index of the calendar for which to retrieve the text. The
	///            control can display up to 12 calendars at once.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CalendarGridCellText, get_CaretDateText
	virtual HRESULT STDMETHODCALLTYPE get_HeaderText(LONG calendarIndex = 0, BSTR* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c HighlightTodayDate property</em>
	///
	/// Retrieves whether the &ldquo;Today&rdquo; date is highlighted, e. g. by drawing a circle around it.
	/// If set to \c VARIANT_TRUE, it is highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HighlightTodayDate, get_ShowTodayDate
	virtual HRESULT STDMETHODCALLTYPE get_HighlightTodayDate(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c HighlightTodayDate property</em>
	///
	/// Sets whether the &ldquo;Today&rdquo; date is highlighted, e. g. by drawing a circle around it.
	/// If set to \c VARIANT_TRUE, it is highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HighlightTodayDate, put_ShowTodayDate
	virtual HRESULT STDMETHODCALLTYPE put_HighlightTodayDate(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c KeepSelectionOnNavigation property</em>
	///
	/// Retrieves whether selected dates remain selected while the user scrolls to the next or previous
	/// view. If set to \c VARIANT_TRUE, the selection is kept on navigation; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_KeepSelectionOnNavigation, get_MaxSelectedDates, get_MultiSelect, get_ScrollRate,
	///     GetSelection, SetSelection
	virtual HRESULT STDMETHODCALLTYPE get_KeepSelectionOnNavigation(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c KeepSelectionOnNavigation property</em>
	///
	/// Sets whether selected dates remain selected while the user scrolls to the next or previous
	/// view. If set to \c VARIANT_TRUE, the selection is kept on navigation; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_KeepSelectionOnNavigation, put_MaxSelectedDates, put_MultiSelect, put_ScrollRate,
	///     GetSelection, SetSelection
	virtual HRESULT STDMETHODCALLTYPE put_KeepSelectionOnNavigation(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c LastDate property</em>
	///
	/// Retrieves the specified calendar's last date that belongs to the month it is displaying.
	///
	/// \param[in] rowIndex The zero-based row index for which to retrieve the date. If set to -1, the last
	///            date of the whole calendar is retrieved.
	/// \param[in] calendarIndex The zero-based index of the calendar for which to retrieve the date. The
	///            control can display up to 12 calendars at once.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_FirstDate, get_MaxDate
	virtual HRESULT STDMETHODCALLTYPE get_LastDate(LONG rowIndex = -1, LONG calendarIndex = 0, DATE* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c MaxDate property</em>
	///
	/// Retrieves the maximum date accepted by the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MaxDate, get_MinDate, get_LastDate, get_CaretDate
	virtual HRESULT STDMETHODCALLTYPE get_MaxDate(DATE* pValue);
	/// \brief <em>Sets the \c MaxDate property</em>
	///
	/// Sets the maximum date accepted by the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MaxDate, put_MinDate, put_CaretDate
	virtual HRESULT STDMETHODCALLTYPE put_MaxDate(DATE newValue);
	/// \brief <em>Retrieves the current setting of the \c MaxSelectedDates property</em>
	///
	/// Retrieves the maximum number of dates that can be selected at the same time.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c MultiSelect property is set to \c VARIANT_FALSE.
	///
	/// \sa put_MaxSelectedDates, get_MultiSelect, GetSelection, get_KeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE get_MaxSelectedDates(LONG* pValue);
	/// \brief <em>Sets the \c MaxSelectedDates property</em>
	///
	/// Sets the maximum number of dates that can be selected at the same time.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c MultiSelect property is set to \c VARIANT_FALSE.
	///
	/// \sa get_MaxSelectedDates, put_MultiSelect, SetSelection, put_KeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE put_MaxSelectedDates(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MinDate property</em>
	///
	/// Retrieves the minimum date accepted by the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MinDate, get_MaxDate, get_FirstDate, get_CaretDate
	virtual HRESULT STDMETHODCALLTYPE get_MinDate(DATE* pValue);
	/// \brief <em>Sets the \c MinDate property</em>
	///
	/// Sets the minimum date accepted by the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MinDate, put_MaxDate, put_CaretDate
	virtual HRESULT STDMETHODCALLTYPE put_MinDate(DATE newValue);
	/// \brief <em>Retrieves the current setting of the \c MonthBackColor property</em>
	///
	/// Retrieves the background color of the month portion of the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_MonthBackColor, get_BackColor, get_ForeColor, get_TitleBackColor, get_TitleForeColor,
	///     get_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_MonthBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c MonthBackColor property</em>
	///
	/// Sets the background color of the month portion of the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_MonthBackColor, put_BackColor, put_ForeColor, put_TitleBackColor, put_TitleForeColor,
	///     put_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_MonthBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, DTCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, DTCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, DTCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, DTCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, DTCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, DTCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MousePointer, get_MouseIcon, DTCtlsLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, DTCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MousePointer, put_MouseIcon, DTCtlsLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, DTCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiSelect property</em>
	///
	/// Retrieves whether multiple dates can be selected at the same time. If set to \c VARIANT_TRUE, the
	/// user may select multiple dates; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MultiSelect, get_MaxSelectedDates, GetSelection, get_KeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE get_MultiSelect(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiSelect property</em>
	///
	/// Sets whether multiple dates can be selected at the same time. If set to \c VARIANT_TRUE, the
	/// user may select multiple dates; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property will destroy and recreate the control window.
	///
	/// \sa get_MultiSelect, put_MaxSelectedDates, SetSelection, put_KeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE put_MultiSelect(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Padding property</em>
	///
	/// Retrieves the distance (in pixels) between the control's content and the frame that the active theme
	/// might draw around the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_Padding, get_UsePadding, get_Appearance, get_BorderStyle
	virtual HRESULT STDMETHODCALLTYPE get_Padding(LONG* pValue);
	/// \brief <em>Sets the \c Padding property</em>
	///
	/// Sets the distance (in pixels) between the control's content and the frame that the active theme
	/// might draw around the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Padding, put_UsePadding, put_Appearance, put_BorderStyle
	virtual HRESULT STDMETHODCALLTYPE put_Padding(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10] or
	/// [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10] or
	/// [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RegisterForOLEDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RegisterForOLEDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, DTCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, DTCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Setting or clearing the \c rtlLayout flag destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, DTCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, DTCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ScrollRate property</em>
	///
	/// Retrieves the scroll rate of the control. The scroll rate is the number of months that the control
	/// moves its display when the user clicks a scroll button.\n
	/// If set to zero, the control scrolls by the number of months it is displaying.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ScrollRate, HitTest, get_KeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE get_ScrollRate(LONG* pValue);
	/// \brief <em>Sets the \c ScrollRate property</em>
	///
	/// Sets the scroll rate of the control. The scroll rate is the number of months that the control
	/// moves its display when the user clicks a scroll button.\n
	/// If set to zero, the control scrolls by the number of months it is displaying.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ScrollRate, HitTest, put_KeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE put_ScrollRate(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowTodayDate property</em>
	///
	/// Retrieves whether the &ldquo;Today&rdquo; date is displayed at the bottom of the control. If set to
	/// \c VARIANT_TRUE, it is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowTodayDate, get_HighlightTodayDate, GetMaxTodayWidth
	virtual HRESULT STDMETHODCALLTYPE get_ShowTodayDate(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowTodayDate property</em>
	///
	/// Sets whether the &ldquo;Today&rdquo; date is displayed at the bottom of the control. If set to
	/// \c VARIANT_TRUE, it is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowTodayDate, put_HighlightTodayDate, GetMaxTodayWidth
	virtual HRESULT STDMETHODCALLTYPE put_ShowTodayDate(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowTrailingDates property</em>
	///
	/// Retrieves whether dates from the previous or next month are displayed to fill up weeks that start or
	/// end in the previous or next month. If set to \c VARIANT_TRUE, such trailing dates are displayed;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_ShowTrailingDates, get_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE get_ShowTrailingDates(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowTrailingDates property</em>
	///
	/// Sets whether dates from the previous or next month are displayed to fill up weeks that start or
	/// end in the previous or next month. If set to \c VARIANT_TRUE, such trailing dates are displayed;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_ShowTrailingDates, put_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE put_ShowTrailingDates(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowWeekNumbers property</em>
	///
	/// Retrieves whether week numbers are displayed to the left of each row of days. Week 1 is defined as
	/// the first week that contains at least four days. If set to \c VARIANT_TRUE, week numbers are
	/// displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowWeekNumbers, get_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE get_ShowWeekNumbers(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowWeekNumbers property</em>
	///
	/// Sets whether week numbers are displayed to the left of each row of days. Week 1 is defined as
	/// the first week that contains at least four days. If set to \c VARIANT_TRUE, week numbers are
	/// displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowWeekNumbers, put_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE put_ShowWeekNumbers(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c StartOfWeek property</em>
	///
	/// Retrieves the day that is handled as the first day of a week. Any of the values defined by the
	/// \c StartOfWeekConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_StartOfWeek, DTCtlsLibU::StartOfWeekConstants
	/// \else
	///   \sa put_StartOfWeek, DTCtlsLibA::StartOfWeekConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StartOfWeek(StartOfWeekConstants* pValue);
	/// \brief <em>Sets the \c StartOfWeek property</em>
	///
	/// Sets the day that is handled as the first day of a week. Any of the values defined by the
	/// \c StartOfWeekConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_StartOfWeek, DTCtlsLibU::StartOfWeekConstants
	/// \else
	///   \sa get_StartOfWeek, DTCtlsLibA::StartOfWeekConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_StartOfWeek(StartOfWeekConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c TitleBackColor property</em>
	///
	/// Retrieves the background color of the title portion of the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_TitleBackColor, get_BackColor, get_ForeColor, get_TitleForeColor, get_MonthBackColor,
	///     get_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_TitleBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c TitleBackColor property</em>
	///
	/// Sets the background color of the title portion of the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_TitleBackColor, put_BackColor, put_ForeColor, put_TitleForeColor, put_MonthBackColor,
	///     put_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_TitleBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c TitleForeColor property</em>
	///
	/// Retrieves the text color of the title portion of the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_TitleForeColor, get_BackColor, get_ForeColor, get_TitleBackColor, get_MonthBackColor,
	///     get_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_TitleForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c TitleForeColor property</em>
	///
	/// Sets the text color of the title portion of the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_TitleForeColor, put_BackColor, put_ForeColor, put_TitleBackColor, put_MonthBackColor,
	///     put_TrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_TitleForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c TodayDate property</em>
	///
	/// Retrieves the date that is displayed as the &ldquo;Today&rdquo; date. If this property is set to zero,
	/// the system default is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_TodayDate, get_ShowTodayDate, get_HighlightTodayDate, get_CaretDate, GetMaxTodayWidth
	virtual HRESULT STDMETHODCALLTYPE get_TodayDate(DATE* pValue);
	/// \brief <em>Sets the \c TodayDate property</em>
	///
	/// Sets the date that is displayed as the &ldquo;Today&rdquo; date. If this property is set to zero,
	/// the system default is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_TodayDate, put_ShowTodayDate, put_HighlightTodayDate, put_CaretDate, GetMaxTodayWidth
	virtual HRESULT STDMETHODCALLTYPE put_TodayDate(DATE newValue);
	/// \brief <em>Retrieves the current setting of the \c TrailingForeColor property</em>
	///
	/// Retrieves the text color of trailing dates.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_TrailingForeColor, get_BackColor, get_ForeColor, get_TitleBackColor, get_TitleForeColor,
	///     get_MonthBackColor
	virtual HRESULT STDMETHODCALLTYPE get_TrailingForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c TrailingForeColor property</em>
	///
	/// Sets the text color of trailing dates.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_TrailingForeColor, put_BackColor, put_ForeColor, put_TitleBackColor, put_TitleForeColor,
	///     put_MonthBackColor
	virtual HRESULT STDMETHODCALLTYPE put_TrailingForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c UsePadding property</em>
	///
	/// Retrieves whether the control applies the padding specified by the \c Padding property. If set to
	/// \c VARIANT_TRUE, the padding is applied; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_UsePadding, get_Padding
	virtual HRESULT STDMETHODCALLTYPE get_UsePadding(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UsePadding property</em>
	///
	/// Sets whether the control applies the padding specified by the \c Padding property. If set to
	/// \c VARIANT_TRUE, the padding is applied; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_UsePadding, put_Padding
	virtual HRESULT STDMETHODCALLTYPE put_UsePadding(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseShortestDayNames property</em>
	///
	/// Retrieves whether the control uses the shortest instead of the short day names for the day of week
	/// column header. If set to \c VARIANT_TRUE, the shortest, otherwise the short names are used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Not all locales seem to define shortest day names. Therefore it may happen that there is
	///          no difference between short day names and shortest day names.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_UseShortestDayNames
	virtual HRESULT STDMETHODCALLTYPE get_UseShortestDayNames(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseShortestDayNames property</em>
	///
	/// Sets whether the control uses the shortest instead of the short day names for the day of week
	/// column header. If set to \c VARIANT_TRUE, the shortest, otherwise the short names are used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Not all locales seem to define shortest day names. Therefore it may happen that there is
	///          no difference between short day names and shortest day names.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_UseShortestDayNames
	virtual HRESULT STDMETHODCALLTYPE put_UseShortestDayNames(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the system's default dialog font or the font specified by the
	/// \c Font property. If set to \c VARIANT_TRUE, the system font; otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_UseSystemFont, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the system's default dialog font or the font specified by the
	/// \c Font property. If set to \c VARIANT_TRUE, the system font; otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c View property</em>
	///
	/// Retrieves the control's view mode. Any of the values defined by the \c ViewConstants enumeration
	/// is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa put_View, Raise_ViewChanged, DTCtlsLibU::ViewConstants
	/// \else
	///   \sa put_View, Raise_ViewChanged, DTCtlsLibA::ViewConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_View(ViewConstants* pValue);
	/// \brief <em>Sets the \c View property</em>
	///
	/// Sets the control's view mode. Any of the values defined by the \c ViewConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa get_View, Raise_ViewChanged, DTCtlsLibU::ViewConstants
	/// \else
	///   \sa get_View, Raise_ViewChanged, DTCtlsLibA::ViewConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_View(ViewConstants newValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Retrieves details about the currently visible range of dates</em>
	///
	/// Retrieves details about the currently visible range of dates. These details include the first and
	/// the last visible date and the number of visible months.
	///
	/// \param[in] includePartialVisibleMonths If set to \c True, trailing dates are included; otherwise
	///            only entirely displayed months are counted.
	/// \param[out] pFirstVisibleDate The first visible date.
	/// \param[out] pLastVisibleDate The last visible date.
	/// \param[out] pVisibleMonths Receives the number of months currently visible.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaretDate
	virtual HRESULT STDMETHODCALLTYPE CountVisibleMonths(VARIANT_BOOL includePartialVisibleMonths = VARIANT_TRUE, DATE* pFirstVisibleDate = NULL, DATE* pLastVisibleDate = NULL, LONG* pVisibleMonths = NULL);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Retrieves the maximum width of the &ldquo;Today&rdquo; string</em>
	///
	/// Retrieves the maximum width of the &ldquo;Today&rdquo; string.
	///
	/// \param[out] pWidth Receives the maximum width of the &ldquo;Today&rdquo; string.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowTodayDate, get_TodayDate
	virtual HRESULT STDMETHODCALLTYPE GetMaxTodayWidth(OLE_XSIZE_PIXELS* pWidth);
	/// \brief <em>Retrieves the control's minimum size required to display a full month</em>
	///
	/// Retrieves the control's minimum size required to display a full month.
	///
	/// \param[out] pWidth The minimum width.
	/// \param[out] pHeight The minimum height.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa MinimizeRectangle
	virtual HRESULT STDMETHODCALLTYPE GetMinimumSize(OLE_XSIZE_PIXELS* pWidth = NULL, OLE_YSIZE_PIXELS* pHeight = NULL);
	/// \brief <em>Retrieves the bounding rectangle of the specified control part</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's upper-left corner) of the
	/// specified control part.
	///
	/// \param[in] controlPart The part of the control for which to retrieve the rectangle. Any of the
	///            values defined by the \c ControlPartConstants enumeration is valid.
	/// \param[in] columnIndex The column index of the grid cell for which to retrieve the rectangle. The
	///            index starts at -1, which identifies the week numbers column. This parameter is ignored
	///            if \c controlPart is not equal to \c cpCalendarCell.
	/// \param[in] rowIndex The row index of the grid cell or row for which to retrieve the rectangle. The
	///            index starts at -1, which identifies the week days row. This parameter is ignored
	///            if \c controlPart is not equal to \c cpCalendarRow or \c cpCalendarCell.
	/// \param[in] calendarIndex The zero-based index of the calendar for which to retrieve the rectangle.
	///            The control can display up to 12 calendars at once. This parameter is ignored
	///            if \c controlPart is equal to \c cpControl, \c cpNextButton, \c cpPreviousButton or
	///            \c cpFooter.
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the control
	///             part's bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the control part's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the control
	///             part's bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the control
	///             part's bounding rectangle relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::ControlPartConstants
	/// \else
	///   \sa DTCtlsLibA::ControlPartConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(ControlPartConstants controlPart, LONG columnIndex = 0, LONG rowIndex = 0, LONG calendarIndex = 0, OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Retrieves the current selection's start and end date</em>
	///
	/// Retrieves the first and the last date of the currently selected range of dates.
	///
	/// \param[out] pSelectionStart The first selected date.
	/// \param[out] pSelectionEnd The last selected date.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method fails if the \c MultiSelect property is set to \c VARIANT_FALSE.
	///
	/// \sa SetSelection, get_CalendarGridCellSelected, get_MultiSelect, get_KeepSelectionOnNavigation,
	///     Raise_SelectionChanged
	virtual HRESULT STDMETHODCALLTYPE GetSelection(DATE* pSelectionStart = NULL, DATE* pSelectionEnd = NULL);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pHitTestDetails Receives a value specifying the exact part of the control the
	///             specified point lies in. Any of the values defined by the \c HitTestConstants
	///             enumeration is valid.
	/// \param[out] pIndexOfHitCalendar <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives
	///             the zero-based index of the calendar that contains the specified point (in case that
	///             multiple calendars are displayed).
	/// \param[out] pIndexOfHitColumn <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives
	///             the zero-based index of the column in the calendar grid that the specified point is
	///             over.
	/// \param[out] pIndexOfHitRow <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives the
	///             zero-based index of the row in the calendar grid that the specified point is over.
	/// \param[out] pCellRectangleLeft <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives
	///             the distance (in pixels) between the control's left border and the left border of the
	///             calendar grid cell that the specified point is over.
	/// \param[out] pCellRectangleTop <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives
	///             the distance (in pixels) between the control's upper border and the upper border of the
	///             calendar grid cell that the specified point is over.
	/// \param[out] pCellRectangleRight <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives
	///             the distance (in pixels) between the control's left border and the right border of the
	///             calendar grid cell that the specified point is over.
	/// \param[out] pCellRectangleBottom <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives
	///             the distance (in pixels) between the control's upper border and the lower border of the
	///             calendar grid cell that the specified point is over.
	/// \param[out] pHitDate Receives the date that the specified point is over, or zero if no date was hit.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::HitTestConstants, HitTest
	/// \else
	///   \sa DTCtlsLibA::HitTestConstants, HitTest
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails = NULL, LONG* pIndexOfHitCalendar = NULL, LONG* pIndexOfHitColumn = NULL, LONG* pIndexOfHitRow = NULL, OLE_XPOS_PIXELS* pCellRectangleLeft = NULL, OLE_YPOS_PIXELS* pCellRectangleTop = NULL, OLE_XPOS_PIXELS* pCellRectangleRight = NULL, OLE_YPOS_PIXELS* pCellRectangleBottom = NULL, DATE* pHitDate = NULL);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Changes the specified rectangle to the smallest size that still fits the same number of calendars</em>
	///
	/// Calculates how many calendars will fit in the specified rectangle, and then sets the rectangle to
	/// the minimum size that it needs to be to fit that number of calendars.
	///
	/// \param[in,out] pLeft On entry, specifies the x-coordinate (in pixels) of the left border of the
	///                rectangle to minimize. On return, it receives the x-coordinate of the left border of
	///                the minimized rectangle.
	/// \param[in,out] pTop On entry, specifies the y-coordinate (in pixels) of the top border of the
	///                rectangle to minimize. On return, it receives the y-coordinate of the top border of
	///                the minimized rectangle.
	/// \param[in,out] pRight On entry, specifies the x-coordinate (in pixels) of the right border of the
	///                rectangle to minimize. On return, it receives the x-coordinate of the right border of
	///                the minimized rectangle.
	/// \param[in,out] pBottom On entry, specifies the y-coordinate (in pixels) of the bottom border of the
	///                rectangle to minimize. On return, it receives the y-coordinate of the bottom border of
	///                the minimized rectangle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa GetMinimumSize, get_CalendarCount
	virtual HRESULT STDMETHODCALLTYPE MinimizeRectangle(OLE_XPOS_PIXELS* pLeft, OLE_YPOS_PIXELS* pTop, OLE_XPOS_PIXELS* pRight, OLE_YPOS_PIXELS* pBottom);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Sets the dates that are displayed using a bold font</em>
	///
	/// Sets the dates that are displayed using a bold font.
	///
	/// \param[in] ppStates An array containing the state for each currently displayed date. If a date's
	///            entry is set to \c VARIANT_TRUE, it is displayed bold; otherwise not.\n
	///            <strong>Note:</strong> To simplify handling, it is assumed that each month has 31 days.
	///            So if the state for April 30th is stored at index 30, the state for May 1st would be
	///            stored at index 32 instead of 31.\n
	///            This array must contain at least \c n * 31 elements where \c n is the number of months
	///            returned by \c CountVisibleMonths(VARIANT_TRUE).
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to a bug in Windows, bold dates won't show up bold if using comctl32.dll version 6.10
	///          on Windows Vista.\n
	///          If the control's view changes or it is closed, any previously set bold states are lost.
	///
	/// \sa CountVisibleMonths, Raise_GetBoldDates
	virtual HRESULT STDMETHODCALLTYPE SetBoldDates(LPSAFEARRAY* ppStates, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Sets the selection's start and end date</em>
	///
	/// Sets the first and the last date of the selected range of dates.
	///
	/// \param[in] selectionStart The first date to select.
	/// \param[in] selectionEnd The last date to select.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method fails if the \c MultiSelect property is set to \c VARIANT_FALSE.
	///
	/// \sa GetSelection, get_CalendarGridCellSelected, get_MultiSelect, get_KeepSelectionOnNavigation,
	///     Raise_SelectionChanged
	virtual HRESULT STDMETHODCALLTYPE SetSelection(DATE selectionStart, DATE selectionEnd);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/056hw3hs.aspx">CComControlBase::OnDraw</a>
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnCreate, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	// \param[in] languageID The locale identifier identifying the language in which name should be
	//            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_GETDLGCODE handler</em>
	///
	/// Will be called if the system needs to know which input messages the control wants to handle.
	/// We use this handler to make the arrow keys work.
	///
	/// \sa PreTranslateAccelerator,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645425.aspx">WM_GETDLGCODE</a>
	LRESULT OnGetDlgCode(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KILLFOCUS handler</em>
	///
	/// Will be called after the control lost the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnSetFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646282.aspx">WM_KILLFOCUS</a>
	LRESULT OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnMButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnSetFocus, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, OnMouseWheel, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnMouseMove, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler mainly to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called after the control gained the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTINGCHANGE handler</em>
	///
	/// Will be called if a system setting was changed.
	/// We use this handler to update our appearance.
	///
	/// \attention This message is posted to top-level windows only, so actually we'll never receive it.
	///
	/// \sa OnThemeChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms725497.aspx">WM_SETTINGCHANGE</a>
	LRESULT OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to update our appearance.
	///
	/// \sa OnSettingChange, Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler for the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFYFORMAT handler</em>
	///
	/// Will be called if the control asks its parent window which format (Unicode/ANSI) the \c WM_NOTIFY
	/// notifications should have.
	/// We use this handler for proper Unicode support.
	///
	/// \sa OnReflectedNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672615.aspx">WM_NOTIFYFORMAT</a>
	LRESULT OnReflectedNotifyFormat(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c MCM_SETCALENDARBORDER handler</em>
	///
	/// Will be called if the control's padding shall be changed.
	/// We use this handler to synchronize the \c UsePadding property.
	///
	/// \sa get_UsePadding,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760993.aspx">MCM_SETCALENDARBORDER</a>
	LRESULT OnSetCalendarBorder(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c MCM_SETCURSEL handler</em>
	///
	/// Will be called if the control's currently selected date shall be changed.
	/// We use this handler to support data binding.
	///
	/// \sa get_CaretDate,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761002.aspx">MCM_SETCURSEL</a>
	LRESULT OnSetCurSel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c MCN_GETDAYSTATE handler</em>
	///
	/// Will be called if the control requests display information about multiple dates from its parent
	/// window.
	/// We use this handler to raise the \c GetBoldDates event.
	///
	/// \sa Raise_GetBoldDates,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760935.aspx">MCN_GETDAYSTATE</a>
	LRESULT OnGetDayStateNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c MCN_SELCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's caret date was changed.
	/// We use this handler to raise the \c SelectionChanged event.
	///
	/// \sa Raise_SelectionChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760937.aspx">MCN_SELCHANGE</a>
	LRESULT OnSelChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c MCN_VIEWCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's view mode was changed.
	/// We use this handler to raise the \c ViewChanged event.
	///
	/// \sa Raise_ViewChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760941.aspx">MCN_VIEWCHANGE</a>
	LRESULT OnViewChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_Click, DTCtlsLibU::_ICalendarEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_Click, DTCtlsLibA::_ICalendarEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the control from
	///                showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_ContextMenu, DTCtlsLibU::_ICalendarEvents::ContextMenu,
	///       Raise_RClick
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_ContextMenu, DTCtlsLibA::_ICalendarEvents::ContextMenu,
	///       Raise_RClick
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu);
	/// \brief <em>Raises the \c DateMouseEnter event</em>
	///
	/// \param[in] date The date that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_DateMouseEnter, DTCtlsLibU::_ICalendarEvents::DateMouseEnter,
	///       Raise_DateMouseLeave, Raise_MouseMove, DTCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_DateMouseEnter, DTCtlsLibA::_ICalendarEvents::DateMouseEnter,
	///       Raise_DateMouseLeave, Raise_MouseMove, DTCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_DateMouseEnter(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c DateMouseLeave event</em>
	///
	/// \param[in] date The date that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_DateMouseLeave, DTCtlsLibU::_ICalendarEvents::DateMouseLeave,
	///       Raise_DateMouseEnter, Raise_MouseMove, DTCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_DateMouseLeave, DTCtlsLibA::_ICalendarEvents::DateMouseLeave,
	///       Raise_DateMouseEnter, Raise_MouseMove, DTCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_DateMouseLeave(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_DblClick, DTCtlsLibU::_ICalendarEvents::DblClick,
	///       Raise_Click, Raise_MDblClick,  Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_DblClick, DTCtlsLibA::_ICalendarEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_DestroyedControlWindow,
	///       DTCtlsLibU::_ICalendarEvents::DestroyedControlWindow, Raise_RecreatedControlWindow, get_hWnd
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_DestroyedControlWindow,
	///       DTCtlsLibA::_ICalendarEvents::DestroyedControlWindow, Raise_RecreatedControlWindow, get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c GetBoldDates event</em>
	///
	/// \param[in] firstDate The first date for which to retrieve the state.
	/// \param[in] numberOfDates The number of dates for which to retrieve the states.
	/// \param[in,out] ppStates An array containing the state for each date. The client may replace this
	///                array.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to a bug in Windows, bold dates won't show up bold if using comctl32.dll version
	///          6.10 on Windows Vista.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_GetBoldDates, DTCtlsLibU::_ICalendarEvents::GetBoldDates,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb760937.aspx">MCN_SELCHANGE</a>,
	///       SetBoldDates
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_GetBoldDates, DTCtlsLibA::_ICalendarEvents::GetBoldDates,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb760937.aspx">MCN_SELCHANGE</a>,
	///       SetBoldDates
	/// \endif
	inline HRESULT Raise_GetBoldDates(DATE firstDate, LONG numberOfDates, LPSAFEARRAY* ppStates);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_KeyDown, DTCtlsLibU::_ICalendarEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_KeyDown, DTCtlsLibA::_ICalendarEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_KeyPress, DTCtlsLibU::_ICalendarEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_KeyPress, DTCtlsLibA::_ICalendarEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp
	/// \endif
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_KeyUp, DTCtlsLibU::_ICalendarEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_KeyUp, DTCtlsLibA::_ICalendarEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MClick, DTCtlsLibU::_ICalendarEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MClick, DTCtlsLibA::_ICalendarEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MDblClick, DTCtlsLibU::_ICalendarEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MDblClick, DTCtlsLibA::_ICalendarEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MouseDown, DTCtlsLibU::_ICalendarEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MouseDown, DTCtlsLibA::_ICalendarEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MouseEnter, DTCtlsLibU::_ICalendarEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_DateMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MouseEnter, DTCtlsLibA::_ICalendarEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_DateMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MouseHover, DTCtlsLibU::_ICalendarEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MouseHover, DTCtlsLibA::_ICalendarEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MouseLeave, DTCtlsLibU::_ICalendarEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_DateMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MouseLeave, DTCtlsLibA::_ICalendarEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_DateMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MouseMove, DTCtlsLibU::_ICalendarEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MouseMove, DTCtlsLibA::_ICalendarEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MouseUp, DTCtlsLibU::_ICalendarEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MouseUp, DTCtlsLibA::_ICalendarEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_MouseWheel, DTCtlsLibU::_ICalendarEvents::MouseWheel,
	///       Raise_MouseMove, DTCtlsLibU::ExtendedMouseButtonConstants, DTCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_MouseWheel, DTCtlsLibA::_ICalendarEvents::MouseWheel,
	///       Raise_MouseMove, DTCtlsLibA::ExtendedMouseButtonConstants, DTCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragDrop, DTCtlsLibU::_ICalendarEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragDrop, DTCtlsLibA::_ICalendarEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data. If \c NULL, the cached data object is used. We use this when
	///            we call this method from other places than \c DragEnter.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragEnter, DTCtlsLibU::_ICalendarEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragEnter, DTCtlsLibA::_ICalendarEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragLeave, DTCtlsLibU::_ICalendarEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragLeave, DTCtlsLibA::_ICalendarEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragMouseMove,
	///       DTCtlsLibU::_ICalendarEvents::OLEDragMouseMove, Raise_OLEDragEnter, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_OLEDragMouseMove,
	///       DTCtlsLibA::_ICalendarEvents::OLEDragMouseMove, Raise_OLEDragEnter, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_RClick, DTCtlsLibU::_ICalendarEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_RClick, DTCtlsLibA::_ICalendarEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_RDblClick, DTCtlsLibU::_ICalendarEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_RDblClick, DTCtlsLibA::_ICalendarEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_RecreatedControlWindow,
	///       DTCtlsLibU::_ICalendarEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_RecreatedControlWindow,
	///       DTCtlsLibA::_ICalendarEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_ResizedControlWindow,
	///       DTCtlsLibU::_ICalendarEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_ResizedControlWindow,
	///       DTCtlsLibA::_ICalendarEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c SelectionChanged event</em>
	///
	/// \param[in] firstSelectedDate The first selected date.
	/// \param[in] lastSelectedDate The last selected date.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event maps directly to the \c MCN_SELCHANGE notification that the \c SysMonthCal32
	///          window is sending whenever the selection changes. This notification is sent twice (and
	///          therefore the event is raised twice), if the user clicks one of the scroll buttons to
	///          view another month.\n
	///          Also \c MCN_SELCHANGE is sent (and the event raised) at regular intervals to detect date
	///          changes.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_SelectionChanged,
	///       DTCtlsLibU::_ICalendarEvents::SelectionChanged, GetSelection, SetSelection, get_CaretDate,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb760937.aspx">MCN_SELCHANGE</a>
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_SelectionChanged,
	///       DTCtlsLibA::_ICalendarEvents::SelectionChanged, GetSelection, SetSelection, get_CaretDate,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb760937.aspx">MCN_SELCHANGE</a>
	/// \endif
	inline HRESULT Raise_SelectionChanged(DATE firstSelectedDate, DATE lastSelectedDate);
	/// \brief <em>Raises the \c ViewChanged event</em>
	///
	/// \param[in] oldView The previous view mode.
	/// \param[in] newView The new view mode.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_ViewChanged, DTCtlsLibU::_ICalendarEvents::ViewChanged,
	///       get_View
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_ViewChanged, DTCtlsLibA::_ICalendarEvents::ViewChanged,
	///       get_View
	/// \endif
	inline HRESULT Raise_ViewChanged(ViewConstants oldView, ViewConstants newView);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_XClick, DTCtlsLibU::_ICalendarEvents::XClick, Raise_XDblClick,
	///       Raise_Click, Raise_MClick, Raise_RClick, DTCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_XClick, DTCtlsLibA::_ICalendarEvents::XClick, Raise_XDblClick,
	///       Raise_Click, Raise_MClick, Raise_RClick, DTCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_ICalendarEvents::Fire_XDblClick, DTCtlsLibU::_ICalendarEvents::XDblClick, Raise_XClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_RDblClick, get_DetectDoubleClicks,
	///       DTCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_ICalendarEvents::Fire_XDblClick, DTCtlsLibA::_ICalendarEvents::XDblClick, Raise_XClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_RDblClick, get_DetectDoubleClicks,
	///       DTCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control by sending messages</em>
	///
	/// Sends \c WM_* and \c DTM_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \else
	///   \sa DTCtlsLibA::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \endif
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the date that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of MCHT_* flags, that holds further details about the control's
	///             part below the specified point.
	///
	/// \return The "hit" date.
	///
	/// \sa HitTest
	DATE HitTest(LONG x, LONG y, __in PUINT pFlags);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// \overload
	///
	/// Retrieves the date that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of MCHT_* flags, that holds further details about the control's
	///             part below the specified point.
	/// \param[out] pIndexOfHitCalendar <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives
	///             the zero-based index of the calendar that contains the specified point (in case that
	///             multiple calendars are displayed).
	/// \param[out] pIndexOfHitColumn <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives the
	///             zero-based index of the column in the calendar grid that the specified point is over.
	/// \param[out] pIndexOfHitRow <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives the
	///             zero-based index of the row in the calendar grid that the specified point is over.
	/// \param[out] pCellRectangle <strong>Comctl32.dll version 6.10 and newer only:</strong> Receives the
	///             bounding rectangle of the calendar grid cell that the specified point is over.
	///
	/// \return The "hit" date.
	///
	/// \sa HitTest
	DATE HitTest(LONG x, LONG y, __in PUINT pFlags, __in PINT pIndexOfHitCalendar, __in PINT pIndexOfHitColumn, __in PINT pIndexOfHitRow, __in LPRECT pCellRectangle);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Auto-scrolls the control</em>
	///
	/// \sa OnTimer, Raise_OLEDragMouseMove, DragDropStatus::AutoScrolling
	void AutoScroll(void);
	/// \brief <em>Resets the day bold states</em>
	///
	/// If the \c MCS_DAYSTATE style is set, for some reason the control starts with random bold states and
	/// doesn't send \c MCN_GETDAYSTATE. Therefore we use a timer and reset the states manually.
	///
	/// \sa OnTimer
	void ResetDayStates(void);


	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<Calendar> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(Calendar* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<Calendar> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<Calendar> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(Calendar* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<Calendar> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c DetectDoubleClicks property's setting</em>
		///
		/// \sa get_DetectDoubleClicks, put_DetectDoubleClicks
		UINT detectDoubleClicks : 1;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c DragScrollTimeBase property's setting</em>
		///
		/// \sa get_DragScrollTimeBase, put_DragScrollTimeBase
		long dragScrollTimeBase;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c ForeColor property's setting</em>
		///
		/// \sa get_ForeColor, put_ForeColor
		OLE_COLOR foreColor;
		/// \brief <em>Holds the \c HighlightTodayDate property's setting</em>
		///
		/// \sa get_HighlightTodayDate, put_HighlightTodayDate
		UINT highlightTodayDate : 1;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c KeepSelectionOnNavigation property's setting</em>
		///
		/// \sa get_KeepSelectionOnNavigation, put_KeepSelectionOnNavigation
		UINT keepSelectionOnNavigation : 1;
		/// \brief <em>Holds the \c MaxDate property's setting</em>
		///
		/// \sa get_MaxDate, put_MaxDate
		DATE maxDate;
		/// \brief <em>Holds the \c MaxSelectedDates property's setting</em>
		///
		/// \sa get_MaxSelectedDates, put_MaxSelectedDates
		long maxSelectedDates;
		/// \brief <em>Holds the \c MinDate property's setting</em>
		///
		/// \sa get_MinDate, put_MinDate
		DATE minDate;
		/// \brief <em>Holds the \c MonthBackColor property's setting</em>
		///
		/// \sa get_MonthBackColor, put_MonthBackColor
		OLE_COLOR monthBackColor;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c MultiSelect property's setting</em>
		///
		/// \sa get_MultiSelect, put_MultiSelect
		UINT multiSelect : 1;
		/// \brief <em>Holds the \c Padding property's setting</em>
		///
		/// \sa get_Padding, put_Padding
		long padding;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		UINT registerForOLEDragDrop : 1;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c ScrollRate property's setting</em>
		///
		/// \sa get_ScrollRate, put_ScrollRate
		long scrollRate;
		/// \brief <em>Holds the \c ShowTodayDate property's setting</em>
		///
		/// \sa get_ShowTodayDate, put_ShowTodayDate
		UINT showTodayDate : 1;
		/// \brief <em>Holds the \c ShowTrailingDates property's setting</em>
		///
		/// \sa get_ShowTrailingDates, put_ShowTrailingDates
		UINT showTrailingDates : 1;
		/// \brief <em>Holds the \c ShowWeekNumbers property's setting</em>
		///
		/// \sa get_ShowWeekNumbers, put_ShowWeekNumbers
		UINT showWeekNumbers : 1;
		/// \brief <em>Holds the \c StartOfWeek property's setting</em>
		///
		/// \sa get_StartOfWeek, put_StartOfWeek
		StartOfWeekConstants startOfWeek;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c TitleBackColor property's setting</em>
		///
		/// \sa get_TitleBackColor, put_TitleBackColor
		OLE_COLOR titleBackColor;
		/// \brief <em>Holds the \c TitleForeColor property's setting</em>
		///
		/// \sa get_TitleForeColor, put_TitleForeColor
		OLE_COLOR titleForeColor;
		/// \brief <em>Holds the \c TrailingForeColor property's setting</em>
		///
		/// \sa get_TrailingForeColor, put_TrailingForeColor
		OLE_COLOR trailingForeColor;
		/// \brief <em>Holds the \c UsePadding property's setting</em>
		///
		/// \sa get_UsePadding, put_UsePadding
		UINT usePadding : 1;
		/// \brief <em>Holds the \c UseShortestDayNames property's setting</em>
		///
		/// \sa get_UseShortestDayNames, put_UseShortestDayNames
		UINT useShortestDayNames : 1;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;
		/// \brief <em>Holds the \c View property's setting</em>
		///
		/// \sa get_View, put_View
		ViewConstants view;

		Properties()
		{
			ResetToDefaults();
		}

		~Properties()
		{
			Release();
		}

		/// \brief <em>Prepares the object for destruction</em>
		void Release(void)
		{
			font.Release();
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			appearance = a2D;
			backColor = 0x80000000 | COLOR_WINDOW;
			borderStyle = bsNone;
			detectDoubleClicks = FALSE;
			disabledEvents = static_cast<DisabledEventsConstants>(deMouseEvents | deClickEvents | deKeyboardEvents);
			dontRedraw = FALSE;
			dragScrollTimeBase = -1;
			enabled = TRUE;
			foreColor = 0x80000000 | COLOR_WINDOWTEXT;
			highlightTodayDate = TRUE;
			hoverTime = -1;
			keepSelectionOnNavigation = FALSE;
			//SYSTEMTIME max = {9999, 12, 0, 31, 23, 59, 59, 999};
			SYSTEMTIME max = {9999, 12, 0, 31, 0, 0, 0, 0};
			SystemTimeToVariantTime(&max, &maxDate);
			maxSelectedDates = 7;
			SYSTEMTIME min = {1752, 9, 0, 14, 0, 0, 0, 0};
			SystemTimeToVariantTime(&min, &minDate);
			monthBackColor = 0x80000000 | COLOR_WINDOW;
			mousePointer = mpDefault;
			multiSelect = FALSE;
			padding = 4;
			processContextMenuKeys = TRUE;
			registerForOLEDragDrop = FALSE;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			scrollRate = 0;
			showTodayDate = TRUE;
			showTrailingDates = TRUE;
			showWeekNumbers = FALSE;
			startOfWeek = sowUseLocale;
			supportOLEDragImages = TRUE;
			titleBackColor = 0x80000000 | COLOR_ACTIVECAPTION;
			titleForeColor = 0x80000000 | COLOR_CAPTIONTEXT;
			trailingForeColor = 0x80000000 | COLOR_GRAYTEXT;
			usePadding = FALSE;
			useShortestDayNames = FALSE;
			useSystemFont = TRUE;
			view = vMonth;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, the control has been activated by mouse and needs to be UI-activated by \c OnSetFocus</em>
		///
		/// ATL always UI-activates the control in \c OnMouseActivate. If the control is activated by mouse,
		/// \c WM_SETFOCUS is sent after \c WM_MOUSEACTIVATE, but Visual Basic 6 won't raise the \c Validate
		/// event if the control already is UI-activated when it receives the focus. Therefore we need to delay
		/// UI-activation.
		///
		/// \sa OnMouseActivate, OnSetFocus, OnKillFocus
		UINT uiActivationPending : 1;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;

		Flags()
		{
			uiActivationPending = FALSE;
			usingThemes = FALSE;
		}
	} flags;


	/// \brief <em>Holds the date below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deMouseEvents being set.
	DATE dateUnderMouse;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event already was raised</em>
		///
		/// \sa Raise_MouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event already was raised</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa Raise_MouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		POINT lastPosition;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Holds data and flags related to auto-scrolling</em>
		///
		/// \sa AutoScroll
		struct AutoScrolling
		{
			/// \brief <em>Holds the current speed multiplier used for horizontal auto-scrolling</em>
			LONG currentHScrollVelocity;
			/// \brief <em>Holds the current interval of the auto-scroll timer</em>
			LONG currentTimerInterval;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the left</em>
			DWORD lastScroll_Left;
			/// \brief <em>Holds the last point of time at which the control was auto-scrolled to the right</em>
			DWORD lastScroll_Right;

			AutoScrolling()
			{
				Reset();
			}

			/// \brief <em>Resets all member variables to their defaults</em>
			void Reset(void)
			{
				currentHScrollVelocity = 0;
				currentTimerInterval = 0;
				lastScroll_Left = 0;
				lastScroll_Right = 0;
			}
		} autoScrolling;

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pDropTargetHelper = NULL;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			drop_pDataObject = NULL;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			autoScrolling.Reset();
		}
	} dragDropStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-scroll the control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;
		/// \brief <em>The ID of the timer that is used to initialize the day states after control window recreation</em>
		///
		/// \remarks If the \c ID_REDRAW timer already is set, the \c ID_INITDAYSTATES timer isn't set in order
		///          to save resources. The \c ID_REDRAW timer will execute the tasks, that otherwise would be
		///          executed by the \c ID_INITDAYSTATES timer.
		///
		/// \sa ResetDayStates
		static const UINT_PTR ID_INITDAYSTATES = 14;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
		/// \brief <em>The interval of the timer that is used to initialize the day states after control window recreation</em>
		static const UINT INT_INITDAYSTATES = 10;
	} timers;

	//////////////////////////////////////////////////////////////////////
	/// \name Version information
	///
	//@{
	/// \brief <em>Retrieves whether we're using at least version 6.10 of comctl32.dll</em>
	///
	/// \return \c TRUE if we're using comctl32.dll version 6.10 or higher; otherwise \c FALSE.
	BOOL IsComctl32Version610OrNewer(void);
	//@}
	//////////////////////////////////////////////////////////////////////

private:
};     // Calendar

OBJECT_ENTRY_AUTO(__uuidof(Calendar), Calendar)