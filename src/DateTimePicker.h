//////////////////////////////////////////////////////////////////////
/// \class DateTimePicker
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c SysDateTimePick32</em>
///
/// This class superclasses \c SysDateTimePick32 and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo \c IMEFlags is the name of a struct as well as a variable.
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa DTCtlsLibU::IDateTimePicker
/// \else
///   \sa DTCtlsLibA::IDateTimePicker
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "DTCtlsU.h"
#else
	#include "DTCtlsA.h"
#endif
#include "_IDateTimePickerEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "CommonProperties.h"
#include "StringProperties.h"
#include "TargetOLEDataObject.h"


class ATL_NO_VTABLE DateTimePicker : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IDateTimePicker, &IID_IDateTimePicker, &LIBID_DTCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IDateTimePicker, &IID_IDateTimePicker, &LIBID_DTCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<DateTimePicker>,
    public IOleControlImpl<DateTimePicker>,
    public IOleObjectImpl<DateTimePicker>,
    public IOleInPlaceActiveObjectImpl<DateTimePicker>,
    public IViewObjectExImpl<DateTimePicker>,
    public IOleInPlaceObjectWindowlessImpl<DateTimePicker>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<DateTimePicker>,
    public Proxy_IDateTimePickerEvents<DateTimePicker>,
    public IPersistStorageImpl<DateTimePicker>,
    public IPersistPropertyBagImpl<DateTimePicker>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<DateTimePicker>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_DateTimePicker, &__uuidof(_IDateTimePickerEvents), &LIBID_DTCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_DateTimePicker, &__uuidof(_IDateTimePickerEvents), &LIBID_DTCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<DateTimePicker>,
    public CComCoClass<DateTimePicker, &CLSID_DateTimePicker>,
    public CComControl<DateTimePicker>,
    public IPerPropertyBrowsingImpl<DateTimePicker>,
    public IDropTarget,
    public ICategorizeProperties,
    public ICreditsProvider
{
public:
	/// \brief <em>The contained drop-down calendar control</em>
	CContainedWindow containedCalendar;
	/// \brief <em>The contained edit control</em>
	CContainedWindow containedEdit;

	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	DateTimePicker();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_IMEMODE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_DATETIMEPICKER)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("DateTimePickerU"), DATETIMEPICK_CLASSW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("DateTimePickerA"), DATETIMEPICK_CLASSA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(DateTimePicker)
			COM_INTERFACE_ENTRY(IDateTimePicker)
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

		BEGIN_PROP_MAP(DateTimePicker)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AllowNullSelection", DISPID_DTP_ALLOWNULLSELECTION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_DTP_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_DTP_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CalendarBackColor", DISPID_DTP_CALENDARBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("CalendarDragScrollTimeBase", DISPID_DTP_CALENDARDRAGSCROLLTIMEBASE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CalendarFont", DISPID_DTP_CALENDARFONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("CalendarForeColor", DISPID_DTP_CALENDARFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("CalendarHighlightTodayDate", DISPID_DTP_CALENDARHIGHLIGHTTODAYDATE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CalendarHoverTime", DISPID_DTP_CALENDARHOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("CalendarKeepSelectionOnNavigation", DISPID_DTP_CALENDARKEEPSELECTIONONNAVIGATION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CalendarMonthBackColor", DISPID_DTP_CALENDARMONTHBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("CalendarShowTodayDate", DISPID_DTP_CALENDARSHOWTODAYDATE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CalendarShowTrailingDates", DISPID_DTP_CALENDARSHOWTRAILINGDATES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CalendarShowWeekNumbers", DISPID_DTP_CALENDARSHOWWEEKNUMBERS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CalendarTitleBackColor", DISPID_DTP_CALENDARTITLEBACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("CalendarTitleForeColor", DISPID_DTP_CALENDARTITLEFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("CalendarTrailingForeColor", DISPID_DTP_CALENDARTRAILINGFORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("CalendarUseShortestDayNames", DISPID_DTP_CALENDARUSESHORTESTDAYNAMES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CalendarUseSystemFont", DISPID_DTP_CALENDARUSESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("CalendarView", DISPID_DTP_CALENDARVIEW, CLSID_NULL, VT_I4)
			//PROP_ENTRY_TYPE("CustomDateFormat", DISPID_DTP_CUSTOMDATEFORMAT, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("DateFormat", DISPID_DTP_DATEFORMAT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DateSelected", DISPID_DTP_DATESELECTED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DetectDoubleClicks", DISPID_DTP_DETECTDOUBLECLICKS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_DTP_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_DTP_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragDropDownTime", DISPID_DTP_DRAGDROPDOWNTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DropDownAlignment", DISPID_DTP_DROPDOWNALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_DTP_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_DTP_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("HoverTime", DISPID_DTP_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("IMEMode", DISPID_DTP_IMEMODE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaxDate", DISPID_DTP_MAXDATE, CLSID_NULL, VT_DATE)
			PROP_ENTRY_TYPE("MinDate", DISPID_DTP_MINDATE, CLSID_NULL, VT_DATE)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_DTP_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_DTP_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_DTP_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_DTP_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_DTP_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("StartOfWeek", DISPID_DTP_STARTOFWEEK, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Style", DISPID_DTP_STYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_DTP_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_DTP_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(DateTimePicker)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IDateTimePickerEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(DateTimePicker)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_INPUTLANGCHANGE, OnInputLangChange)
			INDEXED_MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnMouseWheel)
			INDEXED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEWHEEL, OnMouseWheel)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			MESSAGE_HANDLER(WM_PARENTNOTIFY, OnParentNotify)
			MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

			MESSAGE_HANDLER(DTM_GETMCSTYLE, OnGetMCStyle)
			MESSAGE_HANDLER(DTM_SETFORMATA, OnSetFormat)
			MESSAGE_HANDLER(DTM_SETFORMATW, OnSetFormat)
			MESSAGE_HANDLER(DTM_SETMCFONT, OnSetMCFont)
			MESSAGE_HANDLER(DTM_SETMCSTYLE, OnSetMCStyle)
			MESSAGE_HANDLER(DTM_SETSYSTEMTIME, OnSetSystemTime)

			REFLECTED_NOTIFY_CODE_HANDLER(NM_SETFOCUS, OnSetFocusNotification)

			REFLECTED_NOTIFY_CODE_HANDLER(DTN_CLOSEUP, OnCloseUpNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_DATETIMECHANGE, OnDateTimeChangeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_DROPDOWN, OnDropDownNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_FORMATA, OnFormatNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_FORMATW, OnFormatNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_FORMATQUERYA, OnFormatQueryNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_FORMATQUERYW, OnFormatQueryNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_USERSTRINGA, OnUserStringNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_USERSTRINGW, OnUserStringNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_WMKEYDOWNA, OnWMKeyDownNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(DTN_WMKEYDOWNW, OnWMKeyDownNotification)

			NOTIFY_CODE_HANDLER(MCN_GETDAYSTATE, OnCalendarGetDayStateNotification)
			NOTIFY_CODE_HANDLER(MCN_VIEWCHANGE, OnCalendarViewChangeNotification)

			CHAIN_MSG_MAP(CComControl<DateTimePicker>)
			ALT_MSG_MAP(1)
			MESSAGE_HANDLER(WM_CHAR, OnCalendarChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnCalendarContextMenu)
			MESSAGE_HANDLER(WM_DESTROY, OnCalendarDestroy)
			MESSAGE_HANDLER(WM_KEYDOWN, OnCalendarKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnCalendarKeyUp)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnCalendarLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnCalendarLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnCalendarMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnCalendarMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnCalendarMouseHover)
			MESSAGE_HANDLER(WM_MOUSEHWHEEL, OnCalendarMouseWheel)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnCalendarMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnCalendarMouseMove)
			MESSAGE_HANDLER(WM_MOUSEWHEEL, OnCalendarMouseWheel)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnCalendarRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnCalendarRButtonUp)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnCalendarKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnCalendarKeyUp)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnCalendarXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnCalendarXButtonUp)
			ALT_MSG_MAP(2)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_DESTROY, OnEditDestroy)
			INDEXED_MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			INDEXED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			INDEXED_MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
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
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the control should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties);
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
	/// \name Implementation of IDateTimePicker
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AllowNullSelection property</em>
	///
	/// Retrieves whether a checkbox is displayed allowing to have no date currently selected. If set to
	/// \c VARIANT_TRUE, the checkbox is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowNullSelection, get_DateSelected, get_CurrentDate
	virtual HRESULT STDMETHODCALLTYPE get_AllowNullSelection(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowNullSelection property</em>
	///
	/// Sets whether a checkbox is displayed allowing to have no date currently selected. If set to
	/// \c VARIANT_TRUE, the checkbox is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property will destroy and recreate the control window.
	///
	/// \sa get_AllowNullSelection, put_DateSelected, put_CurrentDate
	virtual HRESULT STDMETHODCALLTYPE put_AllowNullSelection(VARIANT_BOOL newValue);
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
	///   \sa put_Appearance, get_BorderStyle, DTCtlsLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, DTCtlsLibA::AppearanceConstants
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
	///   \sa get_Appearance, put_BorderStyle, DTCtlsLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, DTCtlsLibA::AppearanceConstants
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
	///   \sa put_BorderStyle, get_Appearance, DTCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, DTCtlsLibA::BorderStyleConstants
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
	///   \sa get_BorderStyle, put_Appearance, DTCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, DTCtlsLibA::BorderStyleConstants
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
	/// \brief <em>Retrieves the current setting of the \c CalendarBackColor property</em>
	///
	/// Retrieves the drop-down calendar control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CalendarBackColor, get_CalendarForeColor, get_CalendarTitleBackColor,
	///     get_CalendarTitleForeColor, get_CalendarMonthBackColor, get_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_CalendarBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c CalendarBackColor property</em>
	///
	/// Sets the drop-down calendar control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CalendarBackColor, put_CalendarForeColor, put_CalendarTitleBackColor,
	///     put_CalendarTitleForeColor, put_CalendarMonthBackColor, put_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_CalendarBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarCaretDateText property</em>
	///
	/// Retrieves the text representing the drop-down calendar's caret date. The string contains the
	/// Unicode character 8206 before each part of the text, e. g. &ldquo;?January, ?2009&rdquo; (? stands
	/// for character 8206). This marker can be used to extract specific parts from the text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property fails if the \c hWndCalendar property is set to 0.\n
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CalendarGridCellText, get_CalendarHeaderText, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE get_CalendarCaretDateText(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarDragScrollTimeBase property</em>
	///
	/// Retrieves the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling of the drop-down calendar control during a drag'n'drop operation.
	/// If set to 0, auto-scrolling is disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CalendarDragScrollTimeBase, get_RegisterForOLEDragDrop, get_DragDropDownTime,
	///     Raise_CalendarOLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_CalendarDragScrollTimeBase(LONG* pValue);
	/// \brief <em>Sets the \c CalendarDragScrollTimeBase property</em>
	///
	/// Sets the period of time (in milliseconds) that is used as the time-base to calculate the
	/// velocity of auto-scrolling of the drop-down calendar control during a drag'n'drop operation.
	/// If set to 0, auto-scrolling is disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CalendarDragScrollTimeBase, put_RegisterForOLEDragDrop, put_DragDropDownTime,
	///     Raise_CalendarOLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_CalendarDragScrollTimeBase(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarFirstDate property</em>
	///
	/// Retrieves the drop-down calendar's first date that belongs to the month it is displaying.
	///
	/// \param[in] rowIndex The zero-based row index for which to retrieve the date. If set to -1, the first
	///            date of the whole calendar is retrieved.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property fails if the \c hWndCalendar property is set to 0.\n
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CalendarLastDate, get_MinDate, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE get_CalendarFirstDate(LONG rowIndex = -1, DATE* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c CalendarFont property</em>
	///
	/// Retrieves the drop-down calendar control's font. It's used to draw the drop-down calendar control's
	/// content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_CalendarFont, putref_CalendarFont, get_UseSystemFont, get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_CalendarFont(IFontDisp** ppFont);
	/// \brief <em>Sets the \c CalendarFont property</em>
	///
	/// Sets the drop-down calendar control's font. It's used to draw the drop-down calendar control's
	/// content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.\n
	///          Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarFont, putref_CalendarFont, put_UseSystemFont, put_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_CalendarFont(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c CalendarFont property</em>
	///
	/// Sets the drop-down calendar control's font. It's used to draw the drop-down calendar control's
	/// content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarFont, put_CalendarFont, put_UseSystemFont, putref_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_CalendarFont(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c CalendarForeColor property</em>
	///
	/// Retrieves the drop-down calendar control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_CalendarForeColor, get_CalendarBackColor, get_CalendarTitleBackColor,
	///     get_CalendarTitleForeColor, get_CalendarMonthBackColor, get_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_CalendarForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c CalendarForeColor property</em>
	///
	/// Sets the drop-down calendar control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarForeColor, put_CalendarBackColor, put_CalendarTitleBackColor,
	///     put_CalendarTitleForeColor, put_CalendarMonthBackColor, put_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_CalendarForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarGridCellSelected property</em>
	///
	/// Retrieves whether the specified drop-down calendar grid cell is selected. If set to \c VARIANT_TRUE,
	/// the cell is selected; otherwise not.
	///
	/// \param[in] columnIndex The column index of the grid cell for which to retrieve the selection state.
	///            The index starts at -1, which identifies the week numbers column.
	/// \param[in] rowIndex The row index of the grid cell for which to retrieve the selection state. The
	///            index starts at -1, which identifies the week days row.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property fails if the \c hWndCalendar property is set to 0.\n
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CalendarGridCellText, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE get_CalendarGridCellSelected(LONG columnIndex, LONG rowIndex, VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarGridCellText property</em>
	///
	/// Retrieves the text being displayed in the specified drop-down calendar grid cell. The string
	/// contains the Unicode character 8206 before each part of the text, e. g. &ldquo;?January, ?2009&rdquo;
	/// (? stands for character 8206). This marker can be used to extract specific parts from the text.
	///
	/// \param[in] columnIndex The column index of the grid cell for which to retrieve the text. The index
	///            starts at -1, which identifies the week numbers column.
	/// \param[in] rowIndex The row index of the grid cell for which to retrieve the text. The index starts
	///            at -1, which identifies the week days row.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property fails if the \c hWndCalendar property is set to 0.\n
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CalendarHeaderText, get_CalendarCaretDateText, get_CalendarGridCellSelected, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE get_CalendarGridCellText(LONG columnIndex, LONG rowIndex, BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarHeaderText property</em>
	///
	/// Retrieves the text being displayed in the drop-down calendar header, for instance &ldquo;January,
	/// 2009&rdquo;. The string contains the Unicode character 8206 before each part of the text, e. g.
	/// &ldquo;?January, ?2009&rdquo; (? stands for character 8206). This marker can be used to extract
	/// specific parts from the text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property fails if the \c hWndCalendar property is set to 0.\n
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CalendarGridCellText, get_CalendarCaretDateText, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE get_CalendarHeaderText(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarHighlightTodayDate property</em>
	///
	/// Retrieves whether the &ldquo;Today&rdquo; date is highlighted in the drop-down calendar control,
	/// e. g. by drawing a circle around it. If set to \c VARIANT_TRUE, it is highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CalendarHighlightTodayDate, get_CalendarShowTodayDate
	virtual HRESULT STDMETHODCALLTYPE get_CalendarHighlightTodayDate(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CalendarHighlightTodayDate property</em>
	///
	/// Sets whether the &ldquo;Today&rdquo; date is highlighted in the drop-down calendar control,
	/// e. g. by drawing a circle around it. If set to \c VARIANT_TRUE, it is highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CalendarHighlightTodayDate, put_CalendarShowTodayDate
	virtual HRESULT STDMETHODCALLTYPE put_CalendarHighlightTodayDate(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarHoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the drop-down calendar
	/// control's client area before the \c CalendarMouseHover event is fired. If set to -1, the system
	/// hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CalendarHoverTime, get_HoverTime, Raise_CalendarMouseHover
	virtual HRESULT STDMETHODCALLTYPE get_CalendarHoverTime(LONG* pValue);
	/// \brief <em>Sets the \c CalendarHoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the drop-down calendar
	/// control's client area before the \c CalendarMouseHover event is fired. If set to -1, the system
	/// hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CalendarHoverTime, put_HoverTime, Raise_CalendarMouseHover
	virtual HRESULT STDMETHODCALLTYPE put_CalendarHoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarKeepSelectionOnNavigation property</em>
	///
	/// Retrieves whether selected dates remain selected while the user scrolls to the next or previous
	/// view in the drop-down calendar control. If set to \c VARIANT_TRUE, the selection is kept on
	/// navigation; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_CalendarKeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE get_CalendarKeepSelectionOnNavigation(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CalendarKeepSelectionOnNavigation property</em>
	///
	/// Sets whether selected dates remain selected while the user scrolls to the next or previous
	/// view in the drop-down calendar control. If set to \c VARIANT_TRUE, the selection is kept on
	/// navigation; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_CalendarKeepSelectionOnNavigation
	virtual HRESULT STDMETHODCALLTYPE put_CalendarKeepSelectionOnNavigation(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarLastDate property</em>
	///
	/// Retrieves the drop-down calendar's last date that belongs to the month it is displaying.
	///
	/// \param[in] rowIndex The zero-based row index for which to retrieve the date. If set to -1, the last
	///            date of the whole calendar is retrieved.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property fails if the \c hWndCalendar property is set to 0.\n
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_CalendarFirstDate, get_MaxDate, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE get_CalendarLastDate(LONG rowIndex = -1, DATE* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c CalendarMonthBackColor property</em>
	///
	/// Retrieves the background color of the month portion of the drop-down calendar control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_CalendarMonthBackColor, get_CalendarBackColor, get_CalendarForeColor,
	///     get_CalendarTitleBackColor, get_CalendarTitleForeColor, get_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_CalendarMonthBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c CalendarMonthBackColor property</em>
	///
	/// Sets the background color of the month portion of the drop-down calendar control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarMonthBackColor, put_CalendarBackColor, put_CalendarForeColor,
	///     put_CalendarTitleBackColor, put_CalendarTitleForeColor, put_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_CalendarMonthBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarShowTodayDate property</em>
	///
	/// Retrieves whether the &ldquo;Today&rdquo; date is displayed at the bottom of the drop-down calendar
	/// control. If set to \c VARIANT_TRUE, it is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CalendarShowTodayDate, get_CalendarHighlightTodayDate
	virtual HRESULT STDMETHODCALLTYPE get_CalendarShowTodayDate(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CalendarShowTodayDate property</em>
	///
	/// Sets whether the &ldquo;Today&rdquo; date is displayed at the bottom of the drop-down calendar
	/// control. If set to \c VARIANT_TRUE, it is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CalendarShowTodayDate, put_CalendarHighlightTodayDate
	virtual HRESULT STDMETHODCALLTYPE put_CalendarShowTodayDate(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarShowTrailingDates property</em>
	///
	/// Retrieves whether dates from the previous or next month are displayed in the drop-down calendar
	/// control to fill up weeks that start or end in the previous or next month. If set to \c VARIANT_TRUE,
	/// such trailing dates are displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_CalendarShowTrailingDates, get_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE get_CalendarShowTrailingDates(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CalendarShowTrailingDates property</em>
	///
	/// Sets whether dates from the previous or next month are displayed in the drop-down calendar
	/// control to fill up weeks that start or end in the previous or next month. If set to \c VARIANT_TRUE,
	/// such trailing dates are displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_CalendarShowTrailingDates, put_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE put_CalendarShowTrailingDates(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarShowWeekNumbers property</em>
	///
	/// Retrieves whether week numbers are displayed to the left of each row of days in the drop-down
	/// calendar control. Week 1 is defined as the first week that contains at least four days. If set to
	/// \c VARIANT_TRUE, week numbers are displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CalendarShowWeekNumbers, get_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE get_CalendarShowWeekNumbers(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CalendarShowWeekNumbers property</em>
	///
	/// Sets whether week numbers are displayed to the left of each row of days in the drop-down
	/// calendar control. Week 1 is defined as the first week that contains at least four days. If set to
	/// \c VARIANT_TRUE, week numbers are displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CalendarShowWeekNumbers, put_StartOfWeek
	virtual HRESULT STDMETHODCALLTYPE put_CalendarShowWeekNumbers(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarTitleBackColor property</em>
	///
	/// Retrieves the background color of the title portion of the drop-down calendar control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_CalendarTitleBackColor, get_CalendarBackColor, get_CalendarForeColor,
	///     get_CalendarTitleForeColor, get_CalendarMonthBackColor, get_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_CalendarTitleBackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c CalendarTitleBackColor property</em>
	///
	/// Sets the background color of the title portion of the drop-down calendar control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarTitleBackColor, put_CalendarBackColor, put_CalendarForeColor,
	///     put_CalendarTitleForeColor, put_CalendarMonthBackColor, put_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_CalendarTitleBackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarTitleForeColor property</em>
	///
	/// Retrieves the text color of the title portion of the drop-down calendar control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_CalendarTitleForeColor, get_CalendarBackColor, get_CalendarForeColor,
	///     get_CalendarTitleBackColor, get_CalendarMonthBackColor, get_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE get_CalendarTitleForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c CalendarTitleForeColor property</em>
	///
	/// Sets the text color of the title portion of the drop-down calendar control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarTitleForeColor, put_CalendarBackColor, put_CalendarForeColor,
	///     put_CalendarTitleBackColor, put_CalendarMonthBackColor, put_CalendarTrailingForeColor
	virtual HRESULT STDMETHODCALLTYPE put_CalendarTitleForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarTodayDate property</em>
	///
	/// Retrieves the date that is displayed as the &ldquo;Today&rdquo; date in the drop-down calendar
	/// control. If this property is set to zero, the system default is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CalendarTodayDate, get_CalendarShowTodayDate, get_CalendarHighlightTodayDate,
	///     get_CurrentDate, CalendarGetMaxTodayWidth
	virtual HRESULT STDMETHODCALLTYPE get_CalendarTodayDate(DATE* pValue);
	/// \brief <em>Sets the \c CalendarTodayDate property</em>
	///
	/// Sets the date that is displayed as the &ldquo;Today&rdquo; date in the drop-down calendar
	/// control. If this property is set to zero, the system default is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CalendarTodayDate, put_CalendarShowTodayDate, put_CalendarHighlightTodayDate,
	///     put_CurrentDate, CalendarGetMaxTodayWidth
	virtual HRESULT STDMETHODCALLTYPE put_CalendarTodayDate(DATE newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarTrailingForeColor property</em>
	///
	/// Retrieves the text color of trailing dates in the drop-down calendar control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_CalendarTrailingForeColor, get_CalendarBackColor, get_CalendarForeColor,
	///     get_CalendarTitleBackColor, get_CalendarTitleForeColor, get_CalendarMonthBackColor
	virtual HRESULT STDMETHODCALLTYPE get_CalendarTrailingForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c CalendarTrailingForeColor property</em>
	///
	/// Sets the text color of trailing dates in the drop-down calendar control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarTrailingForeColor, put_CalendarBackColor, put_CalendarForeColor,
	///     put_CalendarTitleBackColor, put_CalendarTitleForeColor, put_CalendarMonthBackColor
	virtual HRESULT STDMETHODCALLTYPE put_CalendarTrailingForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarUseShortestDayNames property</em>
	///
	/// Retrieves whether the control uses the shortest instead of the short day names for the day of week
	/// column header in the drop-down calendar control. If set to \c VARIANT_TRUE, the shortest, otherwise
	/// the short names are used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Not all locales seem to define shortest day names. Therefore it may happen that there is
	///          no difference between short day names and shortest day names.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_CalendarUseShortestDayNames
	virtual HRESULT STDMETHODCALLTYPE get_CalendarUseShortestDayNames(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CalendarUseShortestDayNames property</em>
	///
	/// Sets whether the control uses the shortest instead of the short day names for the day of week
	/// column header in the drop-down calendar control. If set to \c VARIANT_TRUE, the shortest, otherwise
	/// the short names are used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Not all locales seem to define shortest day names. Therefore it may happen that there is
	///          no difference between short day names and shortest day names.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_CalendarUseShortestDayNames
	virtual HRESULT STDMETHODCALLTYPE put_CalendarUseShortestDayNames(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarUseSystemFont property</em>
	///
	/// Retrieves whether the drop-down calendar control uses the system's default dialog font or the font
	/// specified by the \c CalendarFont property. If set to \c VARIANT_TRUE, the system font; otherwise the
	/// specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa put_CalendarUseSystemFont, get_CalendarFont, get_UseSystemFont
	virtual HRESULT STDMETHODCALLTYPE get_CalendarUseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CalendarUseSystemFont property</em>
	///
	/// Sets whether the drop-down calendar control uses the system's default dialog font or the font
	/// specified by the \c CalendarFont property. If set to \c VARIANT_TRUE, the system font; otherwise the
	/// specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Starting with comctl32.dll version 6.10, this property isn't supported for themed calendar
	///          controls.
	///
	/// \sa get_CalendarUseSystemFont, put_CalendarFont, putref_CalendarFont, put_UseSystemFont
	virtual HRESULT STDMETHODCALLTYPE put_CalendarUseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c CalendarView property</em>
	///
	/// Retrieves the drop-down calendar control's view mode. Any of the values defined by the
	/// \c ViewConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa put_CalendarView, Raise_CalendarViewChanged, DTCtlsLibU::ViewConstants
	/// \else
	///   \sa put_CalendarView, Raise_CalendarViewChanged, DTCtlsLibA::ViewConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CalendarView(ViewConstants* pValue);
	/// \brief <em>Sets the \c CalendarView property</em>
	///
	/// Sets the drop-down calendar control's view mode. Any of the values defined by the
	/// \c ViewConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa get_CalendarView, Raise_CalendarViewChanged, DTCtlsLibU::ViewConstants
	/// \else
	///   \sa get_CalendarView, Raise_CalendarViewChanged, DTCtlsLibA::ViewConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_CalendarView(ViewConstants newValue);
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
	/// \brief <em>Retrieves the current setting of the \c CheckboxObjectState property</em>
	///
	/// Retrieves the accessibility object state of the checkbox which is displayed if the
	/// \c AllowNullSelection property is set to \c VARIANT_TRUE. For a list of possible object states see
	/// the <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">MSDN article</a>.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowNullSelection, get_DropDownButtonObjectState,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">Accessibility Object State Constants</a>
	virtual HRESULT STDMETHODCALLTYPE get_CheckboxObjectState(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CurrentDate property</em>
	///
	/// Retrieves the currently selected date.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa put_CurrentDate, get_MinDate, get_MaxDate, get_DateFormat, Raise_CurrentDateChanged
	virtual HRESULT STDMETHODCALLTYPE get_CurrentDate(DATE* pValue);
	/// \brief <em>Sets the \c CurrentDate property</em>
	///
	/// Sets the currently selected date.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa get_CurrentDate, put_MinDate, put_MaxDate, put_DateFormat, Raise_CurrentDateChanged
	virtual HRESULT STDMETHODCALLTYPE put_CurrentDate(DATE newValue);
	/// \brief <em>Retrieves the current setting of the \c CustomDateFormat property</em>
	///
	/// Retrieves the format that the currently selected date is displayed in if the \c DateFormat
	/// property is set to \c dfCustom. The string may be a usual date and time formatting string with
	/// the following placeholders:
	/// - \c d The one- or two-digit day.
	/// - \c dd The two-digit day. Single-digit day values are preceded by a zero.
	/// - \c ddd The three-character weekday abbreviation.
	/// - \c dddd The full weekday name.
	/// - \c h The one- or two-digit hour in 12-hour format.
	/// - \c hh The two-digit hour in 12-hour format. Single-digit values are preceded by a zero.
	/// - \c H The one- or two-digit hour in 24-hour format.
	/// - \c HH The two-digit hour in 24-hour format. Single-digit values are preceded by a zero.
	/// - \c m The one- or two-digit minute.
	/// - \c mm The two-digit minute. Single-digit values are preceded by a zero.
	/// - \c M The one- or two-digit month number.
	/// - \c MM The two-digit month number. Single-digit values are preceded by a zero.
	/// - \c MMM The three-character month abbreviation.
	/// - \c MMMM The full month name.
	/// - \c t The one-letter AM/PM abbreviation (that is, AM is displayed as "A").
	/// - \c tt The two-letter AM/PM abbreviation (that is, AM is displayed as "AM").
	/// - \c yy The last two digits of the year (that is, 1996 would be displayed as "96").
	/// - \c yyyy The full year (that is, 1996 would be displayed as "1996").
	/// - One or more \c X. This is a placeholder for a callback field. For each group of \c X the
	///   \c GetCallbackFieldTextSize and \c GetCallbackFieldText events will be raised. The number of \c X
	///   identifies the callback field the events refer to.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Any text other than the placeholders should be put into single quotes, e. g.
	///          &ldquo;'Today is: 'hh':'m':'s dddd MMM dd', 'yyyy&rdquo;.
	///
	/// \if UNICODE
	///   \sa put_CustomDateFormat, get_DateFormat, get_CurrentDate, Raise_CallbackFieldKeyDown,
	///       Raise_GetCallbackFieldTextSize, Raise_GetCallbackFieldText, DTCtlsLibU::DateFormatConstants
	/// \else
	///   \sa put_CustomDateFormat, get_DateFormat, get_CurrentDate, Raise_CallbackFieldKeyDown,
	///       Raise_GetCallbackFieldTextSize, Raise_GetCallbackFieldText, DTCtlsLibA::DateFormatConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_CustomDateFormat(BSTR* pValue);
	/// \brief <em>Sets the \c CustomDateFormat property</em>
	///
	/// Sets the format that the currently selected date is displayed in if the \c DateFormat
	/// property is set to \c dfCustom. The string may be a usual date and time formatting string with
	/// the following placeholders:
	/// - \c d The one- or two-digit day.
	/// - \c dd The two-digit day. Single-digit day values are preceded by a zero.
	/// - \c ddd The three-character weekday abbreviation.
	/// - \c dddd The full weekday name.
	/// - \c h The one- or two-digit hour in 12-hour format.
	/// - \c hh The two-digit hour in 12-hour format. Single-digit values are preceded by a zero.
	/// - \c H The one- or two-digit hour in 24-hour format.
	/// - \c HH The two-digit hour in 24-hour format. Single-digit values are preceded by a zero.
	/// - \c m The one- or two-digit minute.
	/// - \c mm The two-digit minute. Single-digit values are preceded by a zero.
	/// - \c M The one- or two-digit month number.
	/// - \c MM The two-digit month number. Single-digit values are preceded by a zero.
	/// - \c MMM The three-character month abbreviation.
	/// - \c MMMM The full month name.
	/// - \c t The one-letter AM/PM abbreviation (that is, AM is displayed as "A").
	/// - \c tt The two-letter AM/PM abbreviation (that is, AM is displayed as "AM").
	/// - \c yy The last two digits of the year (that is, 1996 would be displayed as "96").
	/// - \c yyyy The full year (that is, 1996 would be displayed as "1996").
	/// - One or more \c X. This is a placeholder for a callback field. For each group of \c X the
	///   \c GetCallbackFieldTextSize and \c GetCallbackFieldText events will be raised. The number of \c X
	///   identifies the callback field the events refer to.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Any text other than the placeholders should be put into single quotes, e. g.
	///          &ldquo;'Today is: 'hh':'m':'s dddd MMM dd', 'yyyy&rdquo;.
	///
	/// \if UNICODE
	///   \sa get_CustomDateFormat, put_DateFormat, put_CurrentDate, Raise_GetCallbackFieldTextSize,
	///       Raise_GetCallbackFieldText, DTCtlsLibU::DateFormatConstants
	/// \else
	///   \sa get_CustomDateFormat, put_DateFormat, put_CurrentDate, Raise_GetCallbackFieldTextSize,
	///       Raise_GetCallbackFieldText, DTCtlsLibA::DateFormatConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_CustomDateFormat(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c DateFormat property</em>
	///
	/// Retrieves the format that the currently selected date is displayed in. Any of the values defined
	/// by the \c DateFormatConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DateFormat, get_CustomDateFormat, get_CurrentDate, get_Style,
	///       DTCtlsLibU::DateFormatConstants
	/// \else
	///   \sa put_DateFormat, get_CustomDateFormat, get_CurrentDate, get_Style,
	///       DTCtlsLibA::DateFormatConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DateFormat(DateFormatConstants* pValue);
	/// \brief <em>Sets the \c DateFormat property</em>
	///
	/// Sets the format that the currently selected date is displayed in. Any of the values defined by
	/// the \c DateFormatConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DateFormat, put_CustomDateFormat, put_CurrentDate, put_Style,
	///       DTCtlsLibU::DateFormatConstants
	/// \else
	///   \sa get_DateFormat, put_CustomDateFormat, put_CurrentDate, put_Style,
	///       DTCtlsLibA::DateFormatConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DateFormat(DateFormatConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DateSelected property</em>
	///
	/// Retrieves whether there's currently a date selected. If the \c AllowNullSelection property is set to
	/// \c VARIANT_TRUE, the user may change the selected date to &ldquo;No selection&rdquo;. If this
	/// property is set to \c VARIANT_FALSE, the selection is set to &ldquo;No selection&rdquo;; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DateSelected, get_CurrentDate, get_AllowNullSelection, Raise_CurrentDateChanged
	virtual HRESULT STDMETHODCALLTYPE get_DateSelected(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DateSelected property</em>
	///
	/// Sets whether there's currently a date selected. If the \c AllowNullSelection property is set to
	/// \c VARIANT_TRUE, the user may change the selected date to &ldquo;No selection&rdquo;. If this
	/// property is set to \c VARIANT_FALSE, the selection is set to &ldquo;No selection&rdquo;; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DateSelected, put_CurrentDate, put_AllowNullSelection, Raise_CurrentDateChanged
	virtual HRESULT STDMETHODCALLTYPE put_DateSelected(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DetectDoubleClicks property</em>
	///
	/// Retrieves whether double clicks are enabled or disabled. If set to \c VARIANT_TRUE, double clicks are
	/// accepted; otherwise all clicks are handled as single clicks.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Enabling double-clicks may lead to accidental double-clicks.
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
	/// \remarks Enabling double-clicks may lead to accidental double-clicks.
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
	/// \brief <em>Retrieves the current setting of the \c DragDropDownTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be placed over the drop-down button
	/// during a drag'n'drop operation before the drop-down calendar control will be opened automatically.
	/// If set to 0, automatic drop-down is disabled. If set to -1, the system's double-click time,
	/// multiplied with 4, is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragDropDownTime, get_RegisterForOLEDragDrop, get_CalendarDragScrollTimeBase,
	///     Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragDropDownTime(LONG* pValue);
	/// \brief <em>Sets the \c DragDropDownTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be placed over the drop-down button
	/// during a drag'n'drop operation before the drop-down calendar control will be opened automatically.
	/// If set to 0, automatic drop-down is disabled. If set to -1, the system's double-click time,
	/// multiplied with 4, is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragDropDownTime, put_RegisterForOLEDragDrop, put_CalendarDragScrollTimeBase,
	///     Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragDropDownTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DropDownAlignment property</em>
	///
	/// Retrieves the alignment of the drop-down calendar control. Any of the values defined by the
	/// \c DropDownAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DropDownAlignment, DTCtlsLibU::DropDownAlignmentConstants
	/// \else
	///   \sa put_DropDownAlignment, DTCtlsLibA::DropDownAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DropDownAlignment(DropDownAlignmentConstants* pValue);
	/// \brief <em>Sets the \c DropDownAlignment property</em>
	///
	/// Sets the alignment of the drop-down calendar control. Any of the values defined by the
	/// \c DropDownAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DropDownAlignment, DTCtlsLibU::DropDownAlignmentConstants
	/// \else
	///   \sa get_DropDownAlignment, DTCtlsLibA::DropDownAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DropDownAlignment(DropDownAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DropDownButtonObjectState property</em>
	///
	/// Retrieves the accessibility object state of the checkbox which is displayed if the \c Style property
	/// is set to \c sDropDown. For a list of possible object states see the
	/// <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">MSDN article</a>.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Style, get_CheckboxObjectState,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">Accessibility Object State Constants</a>
	virtual HRESULT STDMETHODCALLTYPE get_DropDownButtonObjectState(LONG* pValue);
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
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the control's content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont, get_CalendarFont,
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
	/// \remarks The object identified by \c pNewFont is cloned.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont, put_CalendarFont,
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
	/// \sa get_Font, put_Font, put_UseSystemFont, putref_CalendarFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, get_CalendarHoverTime, Raise_MouseHover
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
	/// \sa get_HoverTime, put_CalendarHoverTime, Raise_MouseHover
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
	/// \sa get_hWndCalendar, get_hWndUpDown, get_hWndEdit, Raise_RecreatedControlWindow,
	///     Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndCalendar property</em>
	///
	/// Retrieves the drop-down calendar control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWnd, get_hWndUpDown, get_hWndEdit, Raise_CalendarDropDown, Raise_CalendarCloseUp
	virtual HRESULT STDMETHODCALLTYPE get_hWndCalendar(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndEdit property</em>
	///
	/// Retrieves the window handle of the edit control that is used during free-format edit mode.
	/// Free-format edit mode can be entered by selecting the whole text or pressing F2 if the
	/// \c deParseUserInput flag of the \c DisabledEvents property is cleared.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_hWnd, get_hWndCalendar, get_hWndUpDown, get_DisabledEvents,
	///       DTCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa get_hWnd, get_hWndCalendar, get_hWndUpDown, get_DisabledEvents,
	///       DTCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_hWndEdit(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndCalendar property</em>
	///
	/// Retrieves the drop-down calendar control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_hWnd, get_hWndCalendar, get_hWndEdit, get_Style, DTCtlsLibU::StyleConstants
	/// \else
	///   \sa get_hWnd, get_hWndCalendar, get_hWndEdit, get_Style, DTCtlsLibA::StyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_hWndUpDown(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c IMEMode property</em>
	///
	/// Retrieves the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IMEMode, DTCtlsLibU::IMEModeConstants
	/// \else
	///   \sa put_IMEMode, DTCtlsLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IMEMode(IMEModeConstants* pValue);
	/// \brief <em>Sets the \c IMEMode property</em>
	///
	/// Sets the control's IME mode. IME is a Windows feature making it easy to enter Asian
	/// characters. Any of the values defined by the \c IMEModeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IMEMode, DTCtlsLibU::IMEModeConstants
	/// \else
	///   \sa get_IMEMode, DTCtlsLibA::IMEModeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IMEMode(IMEModeConstants newValue);
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
	/// \brief <em>Retrieves the current setting of the \c MaxDate property</em>
	///
	/// Retrieves the maximum date accepted by the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MaxDate, get_MinDate, get_CurrentDate
	virtual HRESULT STDMETHODCALLTYPE get_MaxDate(DATE* pValue);
	/// \brief <em>Sets the \c MaxDate property</em>
	///
	/// Sets the maximum date accepted by the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MaxDate, put_MinDate, put_CurrentDate
	virtual HRESULT STDMETHODCALLTYPE put_MaxDate(DATE newValue);
	/// \brief <em>Retrieves the current setting of the \c MinDate property</em>
	///
	/// Retrieves the minimum date accepted by the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MinDate, get_MaxDate, get_CurrentDate
	virtual HRESULT STDMETHODCALLTYPE get_MinDate(DATE* pValue);
	/// \brief <em>Sets the \c MinDate property</em>
	///
	/// Sets the minimum date accepted by the control.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MinDate, put_MaxDate, put_CurrentDate
	virtual HRESULT STDMETHODCALLTYPE put_MinDate(DATE newValue);
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
	/// \sa put_RegisterForOLEDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter,
	///     Raise_CalendarOLEDragEnter
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
	/// \sa get_RegisterForOLEDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter,
	///     Raise_CalendarOLEDragEnter
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
	///   \sa put_RightToLeft, get_IMEMode, DTCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, get_IMEMode, DTCtlsLibA::RightToLeftConstants
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
	///   \sa get_RightToLeft, put_IMEMode, DTCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, put_IMEMode, DTCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
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
	/// \brief <em>Retrieves the current setting of the \c Style property</em>
	///
	/// Retrieves which possibilities to edit the currently selected date are provided. Any of the values
	/// defined by the \c StyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property will destroy and recreate the control window.
	///
	/// \if UNICODE
	///   \sa put_Style, get_DateFormat, get_CurrentDate, DTCtlsLibU::StyleConstants
	/// \else
	///   \sa put_Style, get_DateFormat, get_CurrentDate, DTCtlsLibA::StyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Style(StyleConstants* pValue);
	/// \brief <em>Sets the \c Style property</em>
	///
	/// Sets which possibilities to edit the currently selected date are provided. Any of the values
	/// defined by the \c StyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property will destroy and recreate the control window.
	///
	/// \if UNICODE
	///   \sa get_Style, put_DateFormat, put_CurrentDate, DTCtlsLibU::StyleConstants
	/// \else
	///   \sa get_Style, put_DateFormat, put_CurrentDate, DTCtlsLibA::StyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Style(StyleConstants newValue);
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
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseSystemFont, get_Font, get_CalendarUseSystemFont
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font, put_CalendarUseSystemFont
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Value property</em>
	///
	/// Retrieves the currently selected date. If this property is set to \c Empty and the
	/// \c AllowNullSelection property is set to \c VARIANT_TRUE, the selection is set to &ldquo;No
	/// selection&rdquo;.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property combines the \c CurrentDate and \c DateSelected properties. It can be used
	///          for data-binding if the table may contain \c NULL values.
	///
	/// \sa put_Value, get_CurrentDate, get_DateSelected, Raise_CurrentDateChanged
	virtual HRESULT STDMETHODCALLTYPE get_Value(VARIANT* pValue);
	/// \brief <em>Sets the \c Value property</em>
	///
	/// Sets the currently selected date. If this property is set to \c Empty and the
	/// \c AllowNullSelection property is set to \c VARIANT_TRUE, the selection is set to &ldquo;No
	/// selection&rdquo;.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property combines the \c CurrentDate and \c DateSelected properties. It can be used
	///          for data-binding if the table may contain \c NULL values.
	///
	/// \sa get_Value, put_CurrentDate, put_DateSelected, Raise_CurrentDateChanged
	virtual HRESULT STDMETHODCALLTYPE put_Value(VARIANT newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Retrieves the maximum width of the &ldquo;Today&rdquo; string in the drop-down calendar control</em>
	///
	/// Retrieves the maximum width of the &ldquo;Today&rdquo; string in the drop-down calendar control.
	///
	/// \param[out] pWidth Receives the maximum width of the &ldquo;Today&rdquo; string in the drop-down
	///             calendar control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method fails if the \c hWndCalendar property is set to 0.
	///
	/// \sa get_CalendarShowTodayDate, get_CalendarTodayDate, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE CalendarGetMaxTodayWidth(OLE_XSIZE_PIXELS* pWidth);
	/// \brief <em>Retrieves the bounding rectangle of the specified drop-down calendar control part</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the drop-down calendar control's
	/// upper-left corner) of the specified drop-down calendar control part.
	///
	/// \param[in] controlPart The part of the drop-down calendar control for which to retrieve the
	///            rectangle. Any of the values defined by the \c ControlPartConstants enumeration is valid.
	/// \param[in] columnIndex The column index of the grid cell for which to retrieve the rectangle. The
	///            index starts at -1, which identifies the week numbers column. This parameter is ignored
	///            if \c controlPart is not equal to \c cpCalendarCell.
	/// \param[in] rowIndex The row index of the grid cell or row for which to retrieve the rectangle. The
	///            index starts at -1, which identifies the week days row. This parameter is ignored
	///            if \c controlPart is not equal to \c cpCalendarRow or \c cpCalendarCell.
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the drop-down
	///             calendar control part's bounding rectangle relative to the drop-down calendar control's
	///             upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the drop-down
	///             calendar control part's bounding rectangle relative to the drop-down calendar control's
	///             upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the drop-down
	///             calendar control part's bounding rectangle relative to the drop-down calendar control's
	///             upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the drop-down
	///             calendar control part's bounding rectangle relative to the drop-down calendar control's
	///             upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method fails if the \c hWndCalendar property is set to 0.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::ControlPartConstants, get_hWndCalendar
	/// \else
	///   \sa DTCtlsLibA::ControlPartConstants, get_hWndCalendar
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE CalendarGetRectangle(ControlPartConstants controlPart, LONG columnIndex = 0, LONG rowIndex = 0, OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Closes the drop-down calendar control</em>
	///
	/// Closes the drop-down calendar control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OpenDropDownWindow, Raise_CalendarCloseUp, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE CloseDropDownWindow(void);
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
	/// \brief <em>Retrieves the bounding rectangle of the checkbox that is displayed if \c AllowNullSelection is set to \c VARIANT_TRUE</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's upper-left corner) of the
	/// checkbox which is displayed if the \c AllowNullSelection property is set to \c VARIANT_TRUE.
	///
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the checkbox's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the checkbox's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the checkbox's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the checkbox's
	///             bounding rectangle relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_AllowNullSelection, GetDropDownButtonRectangle
	virtual HRESULT STDMETHODCALLTYPE GetCheckboxRectangle(OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Retrieves the bounding rectangle of the drop-down button that is displayed if \c Style is set to \c sDropDown</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's upper-left corner) of the
	/// drop-down button which is displayed if the \c Style property is set to \c sDropDown.
	///
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the button's
	///             bounding rectangle relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Style, GetCheckboxRectangle
	virtual HRESULT STDMETHODCALLTYPE GetDropDownButtonRectangle(OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Retrieves the control's ideal size so that no clipping occurs</em>
	///
	/// Retrieves the control's ideal size so that no clipping occurs.
	///
	/// \param[out] pWidth The ideal width.
	/// \param[out] pHeight The ideal height.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	virtual HRESULT STDMETHODCALLTYPE GetIdealSize(OLE_XSIZE_PIXELS* pWidth = NULL, OLE_YSIZE_PIXELS* pHeight = NULL);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Opens the drop-down calendar control</em>
	///
	/// Opens the drop-down calendar control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CloseDropDownWindow, Raise_CalendarDropDown, get_hWndCalendar
	virtual HRESULT STDMETHODCALLTYPE OpenDropDownWindow(void);
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
	/// Sets the dates that are displayed using a bold font in the drop-down calendar control.
	///
	/// \param[in] ppStates An array containing the state for each currently displayed date. If a date's
	///            entry is set to \c VARIANT_TRUE, it is displayed bold; otherwise not.\n
	///            <strong>Note:</strong> To simplify handling, it is assumed that each month has 31 days.
	///            So if the state for April 30th is stored at index 30, the state for May 1st would be
	///            stored at index 32 instead of 31.\n
	///            This array must contain at least 93 elements.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          Due to a bug in Windows, bold dates won't show up bold if using comctl32.dll version 6.10
	///          on Windows Vista.\n
	///          If the drop-down calendar control's view changes or it is closed, any previously set bold
	///          states are lost.
	///
	/// \sa Raise_GetBoldDates
	virtual HRESULT STDMETHODCALLTYPE SetBoldDates(LPSAFEARRAY* ppStates, VARIANT_BOOL* pSucceeded);
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
	/// \sa CommonProperties, StringProperties,
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
	/// \sa OnKeyDown, OnCalendarChar, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, OnCalendarContextMenu, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
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
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow, OnCalendarDestroy, OnEditDestroy,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_INPUTLANGCHANGE handler</em>
	///
	/// Will be called after an application's input language has been changed.
	/// We use this handler to update the IME mode of the control.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632629.aspx">WM_INPUTLANGCHANGE</a>
	LRESULT OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, OnCalendarKeyDown, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, OnCalendarKeyUp, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
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
	LRESULT OnLButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, OnCalendarLButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, OnCalendarLButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, OnCalendarMButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, OnCalendarMButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
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
	/// \sa Raise_MouseHover, OnCalendarMouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave, OnCalendarMouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(LONG /*index*/, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, OnMouseWheel, OnCalendarMouseMove, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// control's client area.
	/// We use this handler to raise the \c MouseWheel event.
	///
	/// \sa OnMouseMove, OnCalendarMouseWheel, Raise_MouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnMouseWheel(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_PARENTNOTIFY handler</em>
	///
	/// Will be called if the label-edit control is created or destroyed.
	/// We use this handler to raise the \c CreatedEditControlWindow event.
	///
	/// \sa Raise_CreatedEditControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632638.aspx">WM_PARENTNOTIFY</a>
	LRESULT OnParentNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler mainly to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, OnCalendarRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, OnCalendarRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
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
	/// \sa get_Font, OnSetMCFont,
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
	LRESULT OnXButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, OnCalendarXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, OnCalendarXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFYFORMAT handler</em>
	///
	/// Will be called if the control asks its parent window which format (Unicode/ANSI) the \c WM_NOTIFY
	/// notifications should have.
	/// We use this handler for proper Unicode support.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672615.aspx">WM_NOTIFYFORMAT</a>
	LRESULT OnReflectedNotifyFormat(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTM_GETMCSTYLE handler</em>
	///
	/// Will be called if the style of the control's drop-down calendar control shall be retrieved.
	/// We use this handler to emulate the \c DTM_GETMCSTYLE message, which would require comctl32.dll
	/// version 6.10 or newer, for older systems.
	///
	/// \sa OnSetMCStyle, OnDropDownNotification,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761763.aspx">DTM_GETMCSTYLE</a>
	LRESULT OnGetMCStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTM_SETFORMAT handler</em>
	///
	/// Will be called if the control's currently application-defined display format shall be changed.
	/// We use this handler to synchronize the \c CustomDateFormat property.
	///
	/// \sa get_CustomDateFormat,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761771.aspx">DTM_SETFORMAT</a>
	LRESULT OnSetFormat(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTM_SETMCFONT handler</em>
	///
	/// Will be called if the drop-down calendar control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_CalendarFont, OnSetFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761775.aspx">DTM_SETMCFONT</a>
	LRESULT OnSetMCFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTM_SETMCSTYLE handler</em>
	///
	/// Will be called if the style of the control's drop-down calendar control shall be changed.
	/// We use this handler to emulate the \c DTM_SETMCSTYLE message, which would require comctl32.dll
	/// version 6.10 or newer, for older systems.
	///
	/// \sa OnGetMCStyle, OnDropDownNotification,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761778.aspx">DTM_SETMCSTYLE</a>
	LRESULT OnSetMCStyle(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTM_SETSYSTEMTIME handler</em>
	///
	/// Will be called if the control's currently selected date shall be changed.
	/// We use this handler to support data binding.
	///
	/// \sa get_CurrentDate,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761782.aspx">DTM_SETSYSTEMTIME</a>
	LRESULT OnSetSystemTime(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c CalendarKeyPress event.
	///
	/// \sa OnCalendarKeyDown, OnChar, Raise_CalendarKeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnCalendarChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the drop-down calendar control's context menu should be displayed.
	/// We use this handler to raise the \c CalendarContextMenu event.
	///
	/// \sa OnCalendarRButtonDown, OnContextMenu, Raise_CalendarContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnCalendarContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the drop-down calendar control is being destroyed.
	/// We use this handler to raise the \c DestroyedCalendarControlWindow event.
	///
	/// \sa OnCloseUpNotification, OnEditDestroy, OnDestroy, Raise_DestroyedCalendarControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnCalendarDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the drop-down calendar control has the
	/// keyboard focus.
	/// We use this handler to raise the \c CalendarKeyDown event.
	///
	/// \sa OnCalendarKeyUp, OnCalendarChar, OnKeyDown, Raise_CalendarKeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnCalendarKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the drop-down calendar control has the
	/// keyboard focus.
	/// We use this handler to raise the \c CalendarKeyUp event.
	///
	/// \sa OnCalendarKeyDown, OnCalendarChar, OnKeyUp, Raise_CalendarKeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnCalendarKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the drop-down calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseDown event.
	///
	/// \sa OnCalendarMButtonDown, OnCalendarRButtonDown, OnCalendarXButtonDown, OnLButtonDown,
	///     Raise_CalendarMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnCalendarLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the drop-down calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseUp event.
	///
	/// \sa OnCalendarMButtonUp, OnCalendarRButtonUp, OnCalendarXButtonUp, OnLButtonUp,
	///     Raise_CalendarMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnCalendarLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the drop-down calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseDown event.
	///
	/// \sa OnCalendarLButtonDown, OnCalendarRButtonDown, OnCalendarXButtonDown, OnMButtonDown,
	///     Raise_CalendarMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnCalendarMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the drop-down calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseUp event.
	///
	/// \sa OnCalendarLButtonUp, OnCalendarRButtonUp, OnCalendarXButtonUp, OnMButtonUp,
	///     Raise_CalendarMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnCalendarMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the drop-down calendar control's
	/// client area for a previously specified number of milliseconds.
	/// We use this handler to raise the \c CalendarMouseHover event.
	///
	/// \sa Raise_CalendarMouseHover, OnMouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnCalendarMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the drop-down calendar control's
	/// client area.
	/// We use this handler to raise the \c CalendarMouseLeave event.
	///
	/// \sa Raise_CalendarMouseLeave, OnMouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnCalendarMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the drop-down
	/// calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseMove event.
	///
	/// \sa OnCalendarLButtonDown, OnCalendarLButtonUp, OnCalendarMouseWheel, OnMouseMove,
	///     Raise_CalendarMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnCalendarMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEWHEEL and \c WM_MOUSEHWHEEL handler</em>
	///
	/// Will be called if the user rotates the mouse wheel while the mouse cursor is located over the
	/// drop-down calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseWheel event.
	///
	/// \sa OnCalendarMouseMove, OnMouseWheel, Raise_CalendarMouseWheel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645617.aspx">WM_MOUSEWHEEL</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645614.aspx">WM_MOUSEHWHEEL</a>
	LRESULT OnCalendarMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the drop-down calendar control's client area.
	/// We use this handler mainly to raise the \c CalendarMouseDown event.
	///
	/// \sa OnCalendarLButtonDown, OnCalendarMButtonDown, OnCalendarXButtonDown, OnRButtonDown,
	///     Raise_CalendarMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnCalendarRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the drop-down calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseUp event.
	///
	/// \sa OnCalendarLButtonUp, OnCalendarMButtonUp, OnCalendarXButtonUp, OnRButtonUp,
	///     Raise_CalendarMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnCalendarRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the drop-down calendar control's client area.
	/// We use this handler mainly to raise the \c CalendarMouseDown event.
	///
	/// \sa OnCalendarLButtonDown, OnCalendarMButtonDown, OnCalendarRButtonDown, OnXButtonDown,
	///     Raise_CalendarMouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnCalendarXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the drop-down calendar control's client area.
	/// We use this handler to raise the \c CalendarMouseUp event.
	///
	/// \sa OnCalendarLButtonUp, OnCalendarMButtonUp, OnCalendarRButtonUp, OnXButtonUp,
	///     Raise_CalendarMouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnCalendarXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the edit control is being destroyed.
	/// We use this handler to raise the \c DestroyedEditControlWindow event.
	///
	/// \sa OnCalendarDestroy, OnDestroy, Raise_DestroyedEditControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnEditDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c NM_SETFOCUS handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has gained the keyboard
	/// focus.
	/// We use this handler to initialize IME.
	///
	/// \sa get_IMEMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761751.aspx">NM_SETFOCUS (date time)</a>
	LRESULT OnSetFocusNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled);
	/// \brief <em>\c DTN_CLOSEUP handler</em>
	///
	/// Will be called if the control's parent window is notified, that the drop-down calendar control
	/// is about to be destroyed.
	/// We use this handler to raise the \c CalendarCloseUp event.
	///
	/// \sa OnDropDownNotification, Raise_CalendarCloseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761735.aspx">DTN_CLOSEUP</a>
	LRESULT OnCloseUpNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTN_DATETIMECHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's caret date was changed.
	/// We use this handler to raise the \c CurrentDateChanged event.
	///
	/// \sa Raise_CurrentDateChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761737.aspx">DTN_DATETIMECHANGE</a>
	LRESULT OnDateTimeChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTN_DROPDOWN handler</em>
	///
	/// Will be called if the control's parent window is notified, that the drop-down calendar control
	/// is about to be displayed.
	/// We use this handler to raise the \c CalendarDropDown event and to configure the calendar control.
	///
	/// \sa OnCloseUpNotification, Raise_CalendarDropDown, OnSetMCStyle, OnGetMCStyle,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761739.aspx">DTN_DROPDOWN</a>
	LRESULT OnDropDownNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTN_FORMAT handler</em>
	///
	/// Will be called if the control queries the content of a callback field from its parent window.
	/// We use this handler to raise the \c GetCallbackFieldText event.
	///
	/// \sa OnFormatQueryNotification, OnWMKeyDownNotification, Raise_GetCallbackFieldText,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761741.aspx">DTN_FORMAT</a>
	LRESULT OnFormatNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTN_FORMATQUERY handler</em>
	///
	/// Will be called if the control queries the maximum size of a callback field from its parent window.
	/// We use this handler to raise the \c GetCallbackFieldTextSize event.
	///
	/// \sa OnFormatNotification, Raise_GetCallbackFieldTextSize,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761743.aspx">DTN_FORMATQUERY</a>
	LRESULT OnFormatQueryNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTN_USERSTRING handler</em>
	///
	/// Will be called if the control's parent window is queried to parse user input.
	/// We use this handler to raise the \c ParseUserInput event.
	///
	/// \sa Raise_ParseUserInput,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761745.aspx">DTN_USERSTRING</a>
	LRESULT OnUserStringNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c DTN_WMKEYDOWN handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user has pressed a key while a
	/// callback field had the keyboard focus.
	/// We use this handler to raise the \c CallbackFieldKeyDown event.
	///
	/// \sa OnKeyDown, OnFormatNotification, Raise_CallbackFieldKeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb761747.aspx">DTN_WMKEYDOWN</a>
	LRESULT OnWMKeyDownNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c MCN_GETDAYSTATE handler</em>
	///
	/// Will be called if the drop-down calendar control requests display information about multiple
	/// dates from its parent window.
	/// We use this handler to raise the \c GetBoldDates event.
	///
	/// \sa Raise_GetBoldDates,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760935.aspx">MCN_GETDAYSTATE</a>
	LRESULT OnCalendarGetDayStateNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c MCN_VIEWCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the drop-down calendar
	/// control's view mode was changed.
	/// We use this handler to raise the \c CalendarViewChanged event.
	///
	/// \sa Raise_CalendarViewChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760941.aspx">MCN_VIEWCHANGE</a>
	LRESULT OnCalendarViewChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
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
	/// \brief <em>Raises the \c CalendarClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarClick,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarClick, Raise_CalendarMClick, Raise_CalendarRClick,
	///       Raise_CalendarXClick, Raise_Click
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarClick,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarClick, Raise_CalendarMClick, Raise_CalendarRClick,
	///       Raise_CalendarXClick, Raise_Click
	/// \endif
	inline HRESULT Raise_CalendarClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarCloseUp event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarCloseUp,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarCloseUp, Raise_CalendarDropDown,
	///       CloseDropDownWindow, get_hWndCalendar
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarCloseUp,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarCloseUp, Raise_CalendarDropDown,
	///       CloseDropDownWindow, get_hWndCalendar
	/// \endif
	inline HRESULT Raise_CalendarCloseUp(void);
	/// \brief <em>Raises the \c CalendarContextMenu event</em>
	///
	/// \param[in] date The date that the context menu refers to. May be zero.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The part of the drop-down calendar control that the menu's proposed
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the drop-down calendar
	///                control from showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarContextMenu,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarContextMenu, Raise_CalendarRClick,
	///       Raise_ContextMenu
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarContextMenu,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarContextMenu, Raise_CalendarRClick,
	///       Raise_ContextMenu
	/// \endif
	inline HRESULT Raise_CalendarContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu);
	/// \brief <em>Raises the \c CalendarDateMouseEnter event</em>
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarDateMouseEnter,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarDateMouseEnter, Raise_CalendarDateMouseLeave,
	///       Raise_CalendarMouseMove, DTCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarDateMouseEnter,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarDateMouseEnter, Raise_CalendarDateMouseLeave,
	///       Raise_CalendarMouseMove, DTCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_CalendarDateMouseEnter(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c CalendarDateMouseLeave event</em>
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarDateMouseLeave,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarDateMouseLeave, Raise_CalendarDateMouseEnter,
	///       Raise_CalendarMouseMove, DTCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarDateMouseLeave,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarDateMouseLeave, Raise_CalendarDateMouseEnter,
	///       Raise_CalendarMouseMove, DTCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_CalendarDateMouseLeave(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c CalendarDropDown event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarDropDown,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarDropDown, Raise_CalendarCloseUp,
	///       OpenDropDownWindow, get_hWndCalendar
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarDropDown,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarDropDown, Raise_CalendarCloseUp,
	///       OpenDropDownWindow, get_hWndCalendar
	/// \endif
	inline HRESULT Raise_CalendarDropDown(void);
	/// \brief <em>Raises the \c CalendarKeyDown event</em>
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarKeyDown,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarKeyDown, Raise_CalendarKeyUp,
	///       Raise_CalendarKeyPress, Raise_KeyDown
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarKeyDown,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarKeyDown, Raise_CalendarKeyUp,
	///       Raise_CalendarKeyPress, Raise_KeyDown
	/// \endif
	inline HRESULT Raise_CalendarKeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c CalendarKeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarKeyPress,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarKeyPress, Raise_CalendarKeyDown,
	///       Raise_CalendarKeyUp, Raise_KeyPress
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarKeyPress,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarKeyPress, Raise_CalendarKeyDown,
	///       Raise_CalendarKeyUp, Raise_KeyPress
	/// \endif
	inline HRESULT Raise_CalendarKeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c CalendarKeyUp event</em>
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarKeyUp,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarKeyUp, Raise_CalendarKeyDown,
	///       Raise_CalendarKeyPress, Raise_KeyUp
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarKeyUp,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarKeyUp, Raise_CalendarKeyDown,
	///       Raise_CalendarKeyPress, Raise_KeyUp
	/// \endif
	inline HRESULT Raise_CalendarKeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c CalendarMClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMClick,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMClick, Raise_CalendarClick, Raise_CalendarRClick,
	///       Raise_CalendarXClick, Raise_MClick
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMClick,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMClick, Raise_CalendarClick, Raise_CalendarRClick,
	///       Raise_CalendarXClick, Raise_MClick
	/// \endif
	inline HRESULT Raise_CalendarMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarMouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseDown,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseDown, Raise_CalendarMouseUp,
	///       Raise_CalendarClick, Raise_CalendarMClick, Raise_CalendarRClick, Raise_CalendarXClick,
	///       Raise_MouseDown
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseDown,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseDown, Raise_CalendarMouseUp,
	///       Raise_CalendarClick, Raise_CalendarMClick, Raise_CalendarRClick, Raise_CalendarXClick,
	///       Raise_MouseDown
	/// \endif
	inline HRESULT Raise_CalendarMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarMouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseEnter,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseEnter, Raise_CalendarMouseLeave,
	///       Raise_CalendarDateMouseEnter, Raise_CalendarMouseHover, Raise_CalendarMouseMove,
	///       Raise_MouseEnter
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseEnter,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseEnter, Raise_CalendarMouseLeave,
	///       Raise_CalendarDateMouseEnter, Raise_CalendarMouseHover, Raise_CalendarMouseMove,
	///       Raise_MouseEnter
	/// \endif
	inline HRESULT Raise_CalendarMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarMouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseHover,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseHover, Raise_CalendarMouseEnter,
	///       Raise_CalendarMouseLeave, Raise_CalendarMouseMove, Raise_MouseHover, put_CalendarHoverTime
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseHover,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseHover, Raise_CalendarMouseEnter,
	///       Raise_CalendarMouseLeave, Raise_CalendarMouseMove, Raise_MouseHover, put_CalendarHoverTime
	/// \endif
	inline HRESULT Raise_CalendarMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarMouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseLeave,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseLeave, Raise_CalendarMouseEnter,
	///       Raise_CalendarDateMouseLeave, Raise_CalendarMouseHover, Raise_CalendarMouseMove,
	///       Raise_MouseLeave
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseLeave,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseLeave, Raise_CalendarMouseEnter,
	///       Raise_CalendarDateMouseLeave, Raise_CalendarMouseHover, Raise_CalendarMouseMove,
	///       Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_CalendarMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseMove,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseMove, Raise_CalendarMouseEnter,
	///       Raise_CalendarMouseLeave, Raise_CalendarMouseDown, Raise_CalendarMouseUp,
	///       Raise_CalendarMouseWheel, Raise_MouseMove
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseMove,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseMove, Raise_CalendarMouseEnter,
	///       Raise_CalendarMouseLeave, Raise_CalendarMouseDown, Raise_CalendarMouseUp,
	///       Raise_CalendarMouseWheel, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_CalendarMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarMouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseUp,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseUp, Raise_CalendarMouseDown,
	///       Raise_CalendarClick, Raise_CalendarMClick, Raise_CalendarRClick, Raise_CalendarXClick,
	///       Raise_MouseUp
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseUp,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseUp, Raise_CalendarMouseDown,
	///       Raise_CalendarClick, Raise_CalendarMClick, Raise_CalendarRClick, Raise_CalendarXClick,
	///       Raise_MouseUp
	/// \endif
	inline HRESULT Raise_CalendarMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarMouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseWheel,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseWheel, Raise_CalendarMouseMove,
	///       Raise_MouseWheel, DTCtlsLibU::ExtendedMouseButtonConstants, DTCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarMouseWheel,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseWheel, Raise_CalendarMouseMove,
	///       Raise_MouseWheel, DTCtlsLibA::ExtendedMouseButtonConstants, DTCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_CalendarMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
	/// \brief <em>Raises the \c CalendarOLEDragDrop event</em>
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragDrop,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragDrop, Raise_CalendarOLEDragEnter,
	///       Raise_CalendarOLEDragMouseMove, Raise_CalendarOLEDragLeave, Raise_CalendarMouseUp,
	///       Raise_OLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragDrop,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragDrop, Raise_CalendarOLEDragEnter,
	///       Raise_CalendarOLEDragMouseMove, Raise_CalendarOLEDragLeave, Raise_CalendarMouseUp,
	///       Raise_OLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_CalendarOLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c CalendarOLEDragEnter event</em>
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragEnter,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragEnter, Raise_CalendarOLEDragMouseMove,
	///       Raise_CalendarOLEDragLeave, Raise_CalendarOLEDragDrop, Raise_MouseEnter,
	///       Raise_OLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragEnter,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragEnter, Raise_CalendarOLEDragMouseMove,
	///       Raise_CalendarOLEDragLeave, Raise_CalendarOLEDragDrop, Raise_CalendarMouseEnter,
	///       Raise_OLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_CalendarOLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c CalendarOLEDragLeave event</em>
	///
	/// \param[in] fakedLeave If \c FALSE, the method releases the cached data object; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragLeave,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragLeave, Raise_CalendarOLEDragEnter,
	///       Raise_CalendarOLEDragMouseMove, Raise_CalendarOLEDragDrop, Raise_CalendarMouseLeave,
	///       Raise_OLEDragLeave
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragLeave,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragLeave, Raise_CalendarOLEDragEnter,
	///       Raise_CalendarOLEDragMouseMove, Raise_CalendarOLEDragDrop, Raise_CalendarMouseLeave,
	///       Raise_OLEDragLeave
	/// \endif
	inline HRESULT Raise_CalendarOLEDragLeave(BOOL fakedLeave);
	/// \brief <em>Raises the \c CalendarOLEDragMouseMove event</em>
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragMouseMove,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragMouseMove, Raise_CalendarOLEDragEnter,
	///       Raise_CalendarOLEDragLeave, Raise_CalendarOLEDragDrop, Raise_CalendarMouseMove,
	///       Raise_OLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarOLEDragMouseMove,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragMouseMove, Raise_CalendarOLEDragEnter,
	///       Raise_CalendarOLEDragLeave, Raise_CalendarOLEDragDrop, Raise_CalendarMouseMove,
	///       Raise_OLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_CalendarOLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c CalendarRClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarRClick,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarRClick, Raise_CalendarContextMenu,
	///       Raise_CalendarClick, Raise_CalendarMClick, Raise_CalendarXClick, Raise_RClick
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarRClick,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarRClick, Raise_CalendarContextMenu,
	///       Raise_CalendarClick, Raise_CalendarMClick, Raise_CalendarXClick, Raise_RClick
	/// \endif
	inline HRESULT Raise_CalendarRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CalendarViewChanged event</em>
	///
	/// \param[in] oldView The previous view mode.
	/// \param[in] newView The new view mode.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarViewChanged,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarViewChanged, get_CalendarView
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarViewChanged,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarViewChanged, get_CalendarView
	/// \endif
	inline HRESULT Raise_CalendarViewChanged(ViewConstants oldView, ViewConstants newView);
	/// \brief <em>Raises the \c CalendarXClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarXClick,
	///       DTCtlsLibU::_IDateTimePickerEvents::CalendarXClick, Raise_CalendarClick, Raise_CalendarMClick,
	///       Raise_CalendarRClick, Raise_XClick, DTCtlsLibU::ExtendedMouseButtonConstants,
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CalendarXClick,
	///       DTCtlsLibA::_IDateTimePickerEvents::CalendarXClick, Raise_CalendarClick, Raise_CalendarMClick,
	///       Raise_CalendarRClick, Raise_XClick, DTCtlsLibA::ExtendedMouseButtonConstants,
	/// \endif
	inline HRESULT Raise_CalendarXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CallbackFieldKeyDown event</em>
	///
	/// \param[in] callbackField The callback field for which the text is required. The number of \c X
	///            identifies the field.
	/// \param[in] keyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pCurrentDate The currently selected date. The client usually will change it depending
	///                on the pressed key.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CallbackFieldKeyDown,
	///       DTCtlsLibU::_IDateTimePickerEvents::CallbackFieldKeyDown, Raise_KeyDown,
	///       Raise_GetCallbackFieldText
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CallbackFieldKeyDown,
	///       DTCtlsLibA::_IDateTimePickerEvents::CallbackFieldKeyDown, Raise_KeyDown,
	///       Raise_GetCallbackFieldText
	/// \endif
	inline HRESULT Raise_CallbackFieldKeyDown(BSTR callbackField, SHORT keyCode, SHORT shift, DATE* pCurrentDate);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_Click, DTCtlsLibU::_IDateTimePickerEvents::Click,
	///       Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_Click, DTCtlsLibA::_IDateTimePickerEvents::Click,
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_ContextMenu, DTCtlsLibU::_IDateTimePickerEvents::ContextMenu,
	///       Raise_RClick, Raise_CalendarContextMenu
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_ContextMenu, DTCtlsLibA::_IDateTimePickerEvents::ContextMenu,
	///       Raise_RClick, Raise_CalendarContextMenu
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu);
	/// \brief <em>Raises the \c CreatedCalendarControlWindow event</em>
	///
	/// \param[in] hWndCalendar The drop-down calendar control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CreatedCalendarControlWindow,
	///       DTCtlsLibU::_IDateTimePickerEvents::CreatedCalendarControlWindow,
	///       Raise_DestroyedCalendarControlWindow, Raise_CalendarDropDown, get_hWndCalendar
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CreatedCalendarControlWindow,
	///       DTCtlsLibA::_IDateTimePickerEvents::CreatedCalendarControlWindow,
	///       Raise_DestroyedCalendarControlWindow, Raise_CalendarDropDown, get_hWndCalendar
	/// \endif
	inline HRESULT Raise_CreatedCalendarControlWindow(LONG hWndCalendar);
	/// \brief <em>Raises the \c CreatedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CreatedEditControlWindow,
	///       DTCtlsLibU::_IDateTimePickerEvents::CreatedEditControlWindow, Raise_DestroyedEditControlWindow,
	///       get_hWndEdit
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CreatedEditControlWindow,
	///       DTCtlsLibA::_IDateTimePickerEvents::CreatedEditControlWindow, Raise_DestroyedEditControlWindow,
	///       get_hWndEdit
	/// \endif
	inline HRESULT Raise_CreatedEditControlWindow(LONG hWndEdit);
	/// \brief <em>Raises the \c CurrentDateChanged event</em>
	///
	/// \param[in] dateSelected Specifies whether a date or the &ldquo;No selection&rdquo; state has been
	///            selected. If \c VARIANT_TRUE, the date specified by \c selectedDate has been selected;
	///            otherwise the &ldquo;No selection&rdquo; state has been selected.
	/// \param[in] selectedDate The date that has been selected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event maps directly to the \c DTN_DATETIMECHANGE notification that the
	///          \c SysDateTimePick32 window is sending whenever the selection changes. This notification
	///          sometimes is sent twice (and therefore the event may be raised twice).
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_CurrentDateChanged,
	///       DTCtlsLibU::_IDateTimePickerEvents::CurrentDateChanged, get_CurrentDate, get_DateSelected,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb761737.aspx">DTN_DATETIMECHANGE</a>
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_CurrentDateChanged,
	///       DTCtlsLibA::_IDateTimePickerEvents::CurrentDateChanged, get_CurrentDate, get_DateSelected,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb761737.aspx">DTN_DATETIMECHANGE</a>
	/// \endif
	inline HRESULT Raise_CurrentDateChanged(VARIANT_BOOL dateSelected, DATE selectedDate);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_DblClick, DTCtlsLibU::_IDateTimePickerEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_DblClick, DTCtlsLibA::_IDateTimePickerEvents::DblClick,
	///       Raise_Click, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedCalendarControlWindow event</em>
	///
	/// \param[in] hWndCalendar The drop-down calendar control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_DestroyedCalendarControlWindow,
	///       DTCtlsLibU::_IDateTimePickerEvents::DestroyedCalendarControlWindow,
	///       Raise_CreatedCalendarControlWindow, Raise_CalendarCloseUp,
	///       Raise_DestroyedEditControlWindow, get_hWndCalendar
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_DestroyedCalendarControlWindow,
	///       DTCtlsLibA::_IDateTimePickerEvents::DestroyedCalendarControlWindow,
	///       Raise_CreatedCalendarControlWindow, Raise_CalendarCloseUp,
	///       Raise_DestroyedEditControlWindow, get_hWndCalendar
	/// \endif
	inline HRESULT Raise_DestroyedCalendarControlWindow(LONG hWndCalendar);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_DestroyedControlWindow,
	///       DTCtlsLibU::_IDateTimePickerEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_DestroyedControlWindow,
	///       DTCtlsLibA::_IDateTimePickerEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c DestroyedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_DestroyedEditControlWindow,
	///       DTCtlsLibU::_IDateTimePickerEvents::DestroyedEditControlWindow, get_hWndEdit
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_DestroyedEditControlWindow,
	///       DTCtlsLibA::_IDateTimePickerEvents::DestroyedEditControlWindow, get_hWndEdit
	/// \endif
	inline HRESULT Raise_DestroyedEditControlWindow(LONG hWndEdit);
	/// \brief <em>Raises the \c GetBoldDates event</em>
	///
	/// \param[in] firstDate The first date for which to retrieve the state.
	/// \param[in] numberOfDates The number of dates for which to retrieve the states.
	/// \param[in,out] ppStates An array containing the state for each date. The client may replace this
	///                array.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          Due to a bug in Windows, bold dates won't show up bold if using comctl32.dll version
	///          6.10 on Windows Vista.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_GetBoldDates,
	///       DTCtlsLibU::_IDateTimePickerEvents::GetBoldDates, SetBoldDates
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_GetBoldDates,
	///       DTCtlsLibA::_IDateTimePickerEvents::GetBoldDates, SetBoldDates
	/// \endif
	inline HRESULT Raise_GetBoldDates(DATE firstDate, LONG numberOfDates, LPSAFEARRAY* ppStates);
	/// \brief <em>Raises the \c GetCallbackFieldText event</em>
	///
	/// \param[in] callbackField The callback field for which the text is required. The number of \c X
	///            identifies the field.
	/// \param[in] dateToFormat The date to be formatted.
	/// \param[in,out] pTextHeight Receives the text to display in the callback field.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_GetCallbackFieldText,
	///       DTCtlsLibU::_IDateTimePickerEvents::GetCallbackFieldText, Raise_GetCallbackFieldTextSize,
	///       Raise_CallbackFieldKeyDown, get_CustomDateFormat
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_GetCallbackFieldText,
	///       DTCtlsLibA::_IDateTimePickerEvents::GetCallbackFieldText, Raise_GetCallbackFieldTextSize,
	///       Raise_CallbackFieldKeyDown, get_CustomDateFormat
	/// \endif
	inline HRESULT Raise_GetCallbackFieldText(BSTR callbackField, DATE dateToFormat, BSTR* pTextToDisplay);
	/// \brief <em>Raises the \c GetCallbackFieldTextSize event</em>
	///
	/// \param[in] callbackField The callback field for which the size is required. The number of \c X
	///            identifies the field.
	/// \param[in,out] pTextWidth Receives the maximum width of the callback field's content.
	/// \param[in,out] pTextHeight Receives the maximum height of the callback field's content.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_GetCallbackFieldTextSize,
	///       DTCtlsLibU::_IDateTimePickerEvents::GetCallbackFieldTextSize, Raise_GetCallbackFieldText,
	///       get_CustomDateFormat
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_GetCallbackFieldTextSize,
	///       DTCtlsLibA::_IDateTimePickerEvents::GetCallbackFieldTextSize, Raise_GetCallbackFieldText,
	///       get_CustomDateFormat
	/// \endif
	inline HRESULT Raise_GetCallbackFieldTextSize(BSTR callbackField, OLE_XSIZE_PIXELS* pTextWidth, OLE_YSIZE_PIXELS* pTextHeight);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_KeyDown, DTCtlsLibU::_IDateTimePickerEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress, Raise_CalendarKeyDown, Raise_CallbackFieldKeyDown
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_KeyDown, DTCtlsLibA::_IDateTimePickerEvents::KeyDown,
	///       Raise_KeyUp, Raise_KeyPress, Raise_CalendarKeyDown, Raise_CallbackFieldKeyDown
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_KeyPress, DTCtlsLibU::_IDateTimePickerEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp, Raise_CalendarKeyPress
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_KeyPress, DTCtlsLibA::_IDateTimePickerEvents::KeyPress,
	///       Raise_KeyDown, Raise_KeyUp, Raise_CalendarKeyPress
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_KeyUp, DTCtlsLibU::_IDateTimePickerEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress, Raise_CalendarKeyUp
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_KeyUp, DTCtlsLibA::_IDateTimePickerEvents::KeyUp,
	///       Raise_KeyDown, Raise_KeyPress, Raise_CalendarKeyUp
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MClick, DTCtlsLibU::_IDateTimePickerEvents::MClick,
	///       Raise_MDblClick, Raise_Click, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MClick, DTCtlsLibA::_IDateTimePickerEvents::MClick,
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MDblClick, DTCtlsLibU::_IDateTimePickerEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MDblClick, DTCtlsLibA::_IDateTimePickerEvents::MDblClick,
	///       Raise_MClick, Raise_DblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] index The index of the message map that received the message which led to this event.
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseDown, DTCtlsLibU::_IDateTimePickerEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_CalendarMouseDown
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseDown, DTCtlsLibA::_IDateTimePickerEvents::MouseDown,
	///       Raise_MouseUp, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_CalendarMouseDown
	/// \endif
	inline HRESULT Raise_MouseDown(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] index The index of the message map that received the message which led to this event.
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseEnter, DTCtlsLibU::_IDateTimePickerEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, Raise_CalendarMouseEnter
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseEnter, DTCtlsLibA::_IDateTimePickerEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_MouseHover, Raise_MouseMove, Raise_CalendarMouseEnter
	/// \endif
	inline HRESULT Raise_MouseEnter(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseHover, DTCtlsLibU::_IDateTimePickerEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, Raise_CalendarMouseHover, put_HoverTime
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseHover, DTCtlsLibA::_IDateTimePickerEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, Raise_CalendarMouseHover, put_HoverTime
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseLeave, DTCtlsLibU::_IDateTimePickerEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, Raise_CalendarMouseLeave
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseLeave, DTCtlsLibA::_IDateTimePickerEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_MouseHover, Raise_MouseMove, Raise_CalendarMouseLeave
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] index The index of the message map that received the message which led to this event.
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseMove, DTCtlsLibU::_IDateTimePickerEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       Raise_CalendarMouseMove
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseMove, DTCtlsLibA::_IDateTimePickerEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp, Raise_MouseWheel,
	///       Raise_CalendarMouseMove
	/// \endif
	inline HRESULT Raise_MouseMove(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] index The index of the message map that received the message which led to this event.
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseUp, DTCtlsLibU::_IDateTimePickerEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_CalendarMouseUp
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseUp, DTCtlsLibA::_IDateTimePickerEvents::MouseUp,
	///       Raise_MouseDown, Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick, Raise_CalendarMouseUp
	/// \endif
	inline HRESULT Raise_MouseUp(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseWheel event</em>
	///
	/// \param[in] index The index of the message map that received the message which led to this event.
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseWheel, DTCtlsLibU::_IDateTimePickerEvents::MouseWheel,
	///       Raise_MouseMove, Raise_CalendarMouseWheel, DTCtlsLibU::ExtendedMouseButtonConstants,
	///       DTCtlsLibU::ScrollAxisConstants
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_MouseWheel, DTCtlsLibA::_IDateTimePickerEvents::MouseWheel,
	///       Raise_MouseMove, Raise_CalendarMouseWheel, DTCtlsLibA::ExtendedMouseButtonConstants,
	///       DTCtlsLibA::ScrollAxisConstants
	/// \endif
	inline HRESULT Raise_MouseWheel(LONG index, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragDrop, DTCtlsLibU::_IDateTimePickerEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       Raise_CalendarOLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragDrop, DTCtlsLibA::_IDateTimePickerEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       Raise_CalendarOLEDragDrop, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragEnter,
	///       DTCtlsLibU::_IDateTimePickerEvents::OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseEnter, Raise_CalendarOLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragEnter,
	///       DTCtlsLibA::_IDateTimePickerEvents::OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseEnter, Raise_CalendarOLEDragEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \param[in] fakedLeave If \c FALSE, the method releases the cached data object; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragLeave,
	///       DTCtlsLibU::_IDateTimePickerEvents::OLEDragLeave, Raise_OLEDragEnter, Raise_OLEDragMouseMove,
	///       Raise_OLEDragDrop, Raise_MouseLeave, Raise_CalendarOLEDragLeave
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragLeave,
	///       DTCtlsLibA::_IDateTimePickerEvents::OLEDragLeave, Raise_OLEDragEnter, Raise_OLEDragMouseMove,
	///       Raise_OLEDragDrop, Raise_MouseLeave, Raise_CalendarOLEDragLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(BOOL fakedLeave);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragMouseMove,
	///       DTCtlsLibU::_IDateTimePickerEvents::OLEDragMouseMove, Raise_OLEDragEnter, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseMove, Raise_CalendarOLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_OLEDragMouseMove,
	///       DTCtlsLibA::_IDateTimePickerEvents::OLEDragMouseMove, Raise_OLEDragEnter, Raise_OLEDragLeave,
	///       Raise_OLEDragDrop, Raise_MouseMove, Raise_CalendarOLEDragMouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c ParseUserInput event</em>
	///
	/// \param[in] userInput The string that the user entered and that needs to be parsed.
	/// \param[out] pInputDate Must be set to the date that has been extracted from the user input. To
	///             indicate the &ldquo;No Selection&rdquo; state, this parameter can be set to 0.
	/// \param[out] pDateIsValid Specifies whether the entered string is a valid date. If set to
	///             \c VARIANT_TRUE, it is a valid date; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_ParseUserInput,
	///       DTCtlsLibU::_IDateTimePickerEvents::ParseUserInput, get_CurrentDate, get_DateSelected,
	///       get_DisabledEvents
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_ParseUserInput,
	///       DTCtlsLibA::_IDateTimePickerEvents::ParseUserInput, get_CurrentDate, get_DateSelected,
	///       get_DisabledEvents
	/// \endif
	inline HRESULT Raise_ParseUserInput(BSTR userInput, DATE* pInputDate, VARIANT_BOOL* pDateIsValid);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_RClick, DTCtlsLibU::_IDateTimePickerEvents::RClick,
	///       Raise_ContextMenu, Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_RClick, DTCtlsLibA::_IDateTimePickerEvents::RClick,
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_RDblClick, DTCtlsLibU::_IDateTimePickerEvents::RDblClick,
	///       Raise_RClick, Raise_DblClick, Raise_MDblClick, Raise_XDblClick, get_DetectDoubleClicks
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_RDblClick, DTCtlsLibA::_IDateTimePickerEvents::RDblClick,
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_RecreatedControlWindow,
	///       DTCtlsLibU::_IDateTimePickerEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_RecreatedControlWindow,
	///       DTCtlsLibA::_IDateTimePickerEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IDateTimePickerEvents::Fire_ResizedControlWindow,
	///       DTCtlsLibU::_IDateTimePickerEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_ResizedControlWindow,
	///       DTCtlsLibA::_IDateTimePickerEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_XClick, DTCtlsLibU::_IDateTimePickerEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick, Raise_CalendarXClick,
	///       DTCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_XClick, DTCtlsLibA::_IDateTimePickerEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick, Raise_CalendarXClick,
	///       DTCtlsLibA::ExtendedMouseButtonConstants
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
	///   \sa Proxy_IDateTimePickerEvents::Fire_XDblClick, DTCtlsLibU::_IDateTimePickerEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick,
	///       DTCtlsLibU::ExtendedMouseButtonConstants
	/// \else
	///   \sa Proxy_IDateTimePickerEvents::Fire_XDblClick, DTCtlsLibA::_IDateTimePickerEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick,
	///       DTCtlsLibA::ExtendedMouseButtonConstants
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font, ApplyCalendarFont
	void ApplyFont(void);
	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c CalendarFont property.
	///
	/// \sa get_CalendarFont, ApplyFont
	void ApplyCalendarFont(void);

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
	/// \name IME support
	///
	//@{
	/// \brief <em>Retrieves a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is requested.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa SetCurrentIMEContextMode, DTCtlsLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa SetCurrentIMEContextMode, DTCtlsLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetCurrentIMEContextMode(HWND hWnd);
	/// \brief <em>Sets a window's current IME context mode</em>
	///
	/// \param[in] hWnd The window whose IME context mode is set.
	/// \param[in] IMEMode A constant out of the \c IMEModeConstants enumeration specifying the IME
	///            context mode to apply.
	///
	/// \if UNICODE
	///   \sa GetCurrentIMEContextMode, DTCtlsLibU::IMEModeConstants, put_IMEMode
	/// \else
	///   \sa GetCurrentIMEContextMode, DTCtlsLibA::IMEModeConstants, put_IMEMode
	/// \endif
	void SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode);
	/// \brief <em>Retrieves the control's effective IME context mode</em>
	///
	/// Retrieves the IME context mode that is set for the control after resolving recursive modes like
	/// \c imeInherit.
	///
	/// \return A constant out of the \c IMEModeConstants enumeration specifying the IME context mode.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::IMEModeConstants, get_IMEMode
	/// \else
	///   \sa DTCtlsLibA::IMEModeConstants, get_IMEMode
	/// \endif
	IMEModeConstants GetEffectiveIMEMode(void);
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
	/// Retrieves the date that lies below the point ('x'; 'y') in the drop-down calendar control.
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of MCHT_* flags, that holds further details about the control's
	///             part below the specified point.
	///
	/// \return The "hit" date.
	DATE CalendarHitTest(LONG x, LONG y, __in PUINT pFlags);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// \overload
	///
	/// Retrieves the date that lies below the point ('x'; 'y') in the drop-down calendar control.
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
	DATE CalendarHitTest(LONG x, LONG y, __in PUINT pFlags, __in PINT pIndexOfHitCalendar, __in PINT pIndexOfHitColumn, __in PINT pIndexOfHitRow, __in LPRECT pCellRectangle);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Auto-scrolls the drop-down calendar control</em>
	///
	/// \sa OnTimer, Raise_CalendarOLEDragMouseMove, DragDropStatus::AutoScrolling
	void AutoScroll(void);


	/// \brief <em>Holds constants and flags used with IME support</em>
	struct IMEFlags
	{
	protected:
		/// \brief <em>A table of IME modes to use for Chinese input language</em>
		///
		/// \sa GetIMECountryTable, japaneseIMETable, koreanIMETable
		static IMEModeConstants chineseIMETable[10];
		/// \brief <em>A table of IME modes to use for Japanese input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants japaneseIMETable[10];
		/// \brief <em>A table of IME modes to use for Korean input language</em>
		///
		/// \sa GetIMECountryTable, chineseIMETable, koreanIMETable
		static IMEModeConstants koreanIMETable[10];

	public:
		/// \brief <em>The handle of the default IME context</em>
		HIMC hDefaultIMC;

		IMEFlags()
		{
			hDefaultIMC = NULL;
		}

		/// \brief <em>Retrieves a table of IME modes to use for the current keyboard layout</em>
		///
		/// Retrieves a table of IME modes which can be used to map \c IME_CMODE_* constants to
		/// \c IMEModeConstants constants. The table depends on the current keyboard layout.
		///
		/// \param[in,out] table The IME mode table for the currently active keyboard layout.
		///
		/// \if UNICODE
		///   \sa DTCtlsLibU::IMEModeConstants, GetCurrentIMEContextMode
		/// \else
		///   \sa DTCtlsLibA::IMEModeConstants, GetCurrentIMEContextMode
		/// \endif
		static void GetIMECountryTable(IMEModeConstants table[10])
		{
			WORD languageID = LOWORD(GetKeyboardLayout(0));
			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				switch(languageID) {
					case MAKELANGID(LANG_JAPANESE, SUBLANG_DEFAULT):
						CopyMemory(table, japaneseIMETable, sizeof(japaneseIMETable));
						return;
						break;
					case MAKELANGID(LANG_KOREAN, SUBLANG_DEFAULT):
						CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
						return;
						break;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if(languageID <= MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
				if(languageID == MAKELANGID(LANG_KOREAN, SUBLANG_SYS_DEFAULT)) {
					CopyMemory(table, koreanIMETable, sizeof(koreanIMETable));
					return;
				}
				if(languageID == MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_HONGKONG)) {
					CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
					return;
				}
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			if((languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SINGAPORE)) && (languageID != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_MACAU))) {
				table[0] = static_cast<IMEModeConstants>(-10);
				return;
			}

			CopyMemory(table, chineseIMETable, sizeof(chineseIMETable));
		}
	} IMEFlags;

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
			CComObject< PropertyNotifySinkImpl<DateTimePicker> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(DateTimePicker* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<DateTimePicker> >::CreateInstance(&pPropertyNotifySink);
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
			CComObject< PropertyNotifySinkImpl<DateTimePicker> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(DateTimePicker* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<DateTimePicker> >::CreateInstance(&pPropertyNotifySink);
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

		/// \brief <em>Holds the \c AllowNullSelection property's setting</em>
		///
		/// \sa get_AllowNullSelection, put_AllowNullSelection
		UINT allowNullSelection : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c CalendarBackColor property's setting</em>
		///
		/// \sa get_CalendarBackColor, put_CalendarBackColor
		OLE_COLOR calendarBackColor;
		/// \brief <em>Holds the \c CalendarDragScrollTimeBase property's setting</em>
		///
		/// \sa get_CalendarDragScrollTimeBase, put_CalendarDragScrollTimeBase
		long calendarDragScrollTimeBase;
		/// \brief <em>Holds the \c CalendarFont property's settings</em>
		///
		/// \sa get_CalendarFont, put_CalendarFont, putref_CalendarFont
		FontProperty calendarFont;
		/// \brief <em>Holds the \c CalendarForeColor property's setting</em>
		///
		/// \sa get_CalendarForeColor, put_CalendarForeColor
		OLE_COLOR calendarForeColor;
		/// \brief <em>Holds the \c CalendarHighlightTodayDate property's setting</em>
		///
		/// \sa get_CalendarHighlightTodayDate, put_CalendarHighlightTodayDate
		UINT calendarHighlightTodayDate : 1;
		/// \brief <em>Holds the \c CalendarHoverTime property's setting</em>
		///
		/// \sa get_CalendarHoverTime, put_CalendarHoverTime
		long calendarHoverTime;
		/// \brief <em>Holds the \c CalendarKeepSelectionOnNavigation property's setting</em>
		///
		/// \sa get_CalendarKeepSelectionOnNavigation, put_CalendarKeepSelectionOnNavigation
		UINT calendarKeepSelectionOnNavigation : 1;
		/// \brief <em>Holds the \c CalendarMonthBackColor property's setting</em>
		///
		/// \sa get_CalendarMonthBackColor, put_CalendarMonthBackColor
		OLE_COLOR calendarMonthBackColor;
		/// \brief <em>Holds the \c CalendarShowTodayDate property's setting</em>
		///
		/// \sa get_CalendarShowTodayDate, put_CalendarShowTodayDate
		UINT calendarShowTodayDate : 1;
		/// \brief <em>Holds the \c CalendarShowTrailingDates property's setting</em>
		///
		/// \sa get_CalendarShowTrailingDates, put_CalendarShowTrailingDates
		UINT calendarShowTrailingDates : 1;
		/// \brief <em>Holds the \c CalendarShowWeekNumbers property's setting</em>
		///
		/// \sa get_CalendarShowWeekNumbers, put_CalendarShowWeekNumbers
		UINT calendarShowWeekNumbers : 1;
		/// \brief <em>Holds the \c CalendarTitleBackColor property's setting</em>
		///
		/// \sa get_CalendarTitleBackColor, put_CalendarTitleBackColor
		OLE_COLOR calendarTitleBackColor;
		/// \brief <em>Holds the \c CalendarTitleForeColor property's setting</em>
		///
		/// \sa get_CalendarTitleForeColor, put_CalendarTitleForeColor
		OLE_COLOR calendarTitleForeColor;
		/// \brief <em>Holds the \c CalendarTrailingForeColor property's setting</em>
		///
		/// \sa get_CalendarTrailingForeColor, put_CalendarTrailingForeColor
		OLE_COLOR calendarTrailingForeColor;
		/// \brief <em>Holds the \c CalendarUseShortestDayNames property's setting</em>
		///
		/// \sa get_CalendarUseShortestDayNames, put_CalendarUseShortestDayNames
		UINT calendarUseShortestDayNames : 1;
		/// \brief <em>Holds the \c CalendarUseSystemFont property's setting</em>
		///
		/// \sa get_CalendarUseSystemFont, put_CalendarUseSystemFont
		UINT calendarUseSystemFont : 1;
		/// \brief <em>Holds the \c CalendarView property's setting</em>
		///
		/// \sa get_CalendarView, put_CalendarView
		ViewConstants calendarView;
		/// \brief <em>Holds the \c CurrentDate property's setting</em>
		///
		/// \sa get_CurrentDate, put_CurrentDate
		SYSTEMTIME currentDate;
		/// \brief <em>Holds the \c CustomDateFormat property's setting</em>
		///
		/// \sa get_CustomDateFormat, put_CustomDateFormat
		CComBSTR customDateFormat;
		/// \brief <em>Holds the \c DateFormat property's setting</em>
		///
		/// \sa get_DateFormat, put_DateFormat
		DateFormatConstants dateFormat;
		/// \brief <em>Holds the \c DateSelected property's setting</em>
		///
		/// \sa get_DateSelected, put_DateSelected
		UINT dateSelected : 1;
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
		/// \brief <em>Holds the \c DragDropDownTime property's setting</em>
		///
		/// \sa get_DragDropDownTime, put_DragDropDownTime
		long dragDropDownTime;
		/// \brief <em>Holds the \c DropDownAlignment property's setting</em>
		///
		/// \sa get_DropDownAlignment, put_DropDownAlignment
		DropDownAlignmentConstants dropDownAlignment;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c IMEMode property's setting</em>
		///
		/// \sa get_IMEMode, put_IMEMode
		IMEModeConstants IMEMode;
		/// \brief <em>Holds the \c MaxDate property's setting</em>
		///
		/// \sa get_MaxDate, put_MaxDate
		DATE maxDate;
		/// \brief <em>Holds the \c MinDate property's setting</em>
		///
		/// \sa get_MinDate, put_MinDate
		DATE minDate;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
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
		/// \brief <em>Holds the \c StartOfWeek property's setting</em>
		///
		/// \sa get_StartOfWeek, put_StartOfWeek
		StartOfWeekConstants startOfWeek;
		/// \brief <em>Holds the \c Style property's setting</em>
		///
		/// \sa get_Style, put_Style
		StyleConstants style;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;

		/// \brief <em>Holds the style flags of the drop-down calendar control</em>
		///
		/// \sa OnGetMCStyle, OnSetMCStyle
		DWORD calendarStyle;

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
			allowNullSelection = FALSE;
			appearance = aDefault;
			borderStyle = bsNone;
			calendarBackColor = 0x80000000 | COLOR_WINDOW;
			calendarDragScrollTimeBase = -1;
			calendarForeColor = 0x80000000 | COLOR_WINDOWTEXT;
			calendarHighlightTodayDate = TRUE;
			calendarHoverTime = -1;
			calendarKeepSelectionOnNavigation = FALSE;
			calendarMonthBackColor = 0x80000000 | COLOR_WINDOW;
			calendarShowTodayDate = TRUE;
			calendarShowTrailingDates = TRUE;
			calendarShowWeekNumbers = FALSE;
			calendarTitleBackColor = 0x80000000 | COLOR_ACTIVECAPTION;
			calendarTitleForeColor = 0x80000000 | COLOR_CAPTIONTEXT;
			calendarTrailingForeColor = 0x80000000 | COLOR_GRAYTEXT;
			calendarUseShortestDayNames = FALSE;
			calendarUseSystemFont = TRUE;
			calendarView = vMonth;
			GetSystemTime(&currentDate);
			customDateFormat = L"";
			dateFormat = dfShortDate;
			dateSelected = TRUE;
			detectDoubleClicks = TRUE;
			disabledEvents = static_cast<DisabledEventsConstants>(deMouseEvents | deClickEvents | deCalendarMouseEvents | deKeyboardEvents | deCalendarKeyboardEvents | deCalendarClickEvents | deParseUserInput);
			dontRedraw = FALSE;
			dragDropDownTime = -1;
			dropDownAlignment = ddaLeft;
			enabled = TRUE;
			hoverTime = -1;
			IMEMode = imeInherit;
			//SYSTEMTIME max = {9999, 12, 0, 31, 23, 59, 59, 999};
			SYSTEMTIME max = {9999, 12, 0, 31, 0, 0, 0, 0};
			SystemTimeToVariantTime(&max, &maxDate);
			SYSTEMTIME min = {1752, 9, 0, 14, 0, 0, 0, 0};
			SystemTimeToVariantTime(&min, &minDate);
			mousePointer = mpDefault;
			processContextMenuKeys = TRUE;
			registerForOLEDragDrop = FALSE;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			startOfWeek = sowUseLocale;
			style = sDropDown;
			supportOLEDragImages = TRUE;
			useSystemFont = TRUE;

			calendarStyle = 0;
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
		/// \brief <em>If \c TRUE, the \c CurrentDateChanged event isn't raised</em>
		///
		/// \sa Raise_CurrentDateChanged
		UINT noDateChangeEvent : 1;
		/// \brief <em>If \c TRUE, the \c KeyDown and \c KeyUp events won't be raised for the \c ESC key</em>
		///
		/// \remarks This flag will be reset automatically by the \c OnKeyUp message handler.
		///
		/// \sa ignoreEscapeChar, OnKeyDown, OnKeyUp, CloseDropDownWindow
		UINT ignoreEscapeKeyDownAndUp : 1;
		/// \brief <em>If larger than 0, the \c KeyPress event won't be raised for the \c ESC key</em>
		///
		/// \remarks The \c OnChar message handler decrements this value by 1 after applying it.
		///
		/// \sa ignoreEscapeKeyDownAndUp, OnChar, CloseDropDownWindow
		UINT ignoreEscapeChar : 2;
		/// \brief <em>If \c TRUE, the control applies a hack that works around a bug in version 5.x of comctl32.dll</em>
		///
		/// The implementation of \c SysDateTimePick32 displays the drop-down calendar on \c WM_LBUTTONDOWN.
		/// At least in version 5.x of comctl32.dll, \c SysDateTimePick32 enters a message loop after
		/// displaying the drop-down calendar. This message loop is left, if a \c WM_COMMAND message is
		/// dispatched. When leaving the message loop, the drop-down calendar is closed up.\n
		/// This behavior leads to problems if an edit control has the focus and the user clicks onto the
		/// drop-down button of the date time picker. Then \c SysDateTimePick32 will display the drop-down
		/// calendar and enter the message loop and then the edit control sends a \c WM_COMMAND message to
		/// its parent window signalling that it has lost the focus. Since the message loop, that
		/// \c SysDateTimePick32 is in, gets any thread messages, it will also get this \c WM_COMMAND message.
		/// The result is that the first click onto the drop-down button will fail and lead to a still-pressed
		/// button, if previously a control had the focus that issues a \c WM_COMMAND message when it loses the
		/// focus.\n
		/// Our work-around detects (to some degree) this situation in advance and does not raise any of the
		/// events that would be raised normally, until \c WM_LBUTTONUP is received. On \c WM_LBUTTONUP,
		/// another mouse click is faked. This click together with the initial click will be handled as a
		/// double click. Thus it needs to be ignored, too. After that, the \c WM_LBUTTONUP message handler
		/// fakes another mouse click. This time it will cause the drop-down calendar to appear and won't be
		/// suppressed.
		///
		/// \remarks This flag will be reset automatically by the \c OnLButtonUp message handler.
		///
		/// \sa ignoreDblClick, OnLButtonUp, OnLButtonDown
		UINT applyDropDownWorkaround : 1;
		/// \brief <em>If \c TRUE, the \c DblClick event isn't raised</em>
		///
		/// \remarks This flag will be reset automatically by the \c OnLButtonDblClk message handler.
		///
		/// \sa OnLButtonUp, OnLButtonDblClk
		UINT ignoreDblClick : 1;

		Flags()
		{
			uiActivationPending = FALSE;
			usingThemes = FALSE;
			noDateChangeEvent = FALSE;
			ignoreEscapeKeyDownAndUp = FALSE;
			ignoreEscapeChar = 0;
			applyDropDownWorkaround = FALSE;
			ignoreDblClick = FALSE;
		}
	} flags;


	/// \brief <em>Holds the date below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deCalendarMouseEvents being set.
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
	///
	/// \sa mouseStatus_Calendar
	MouseStatus mouseStatus;
	/// \brief <em>Holds the drop-down calendar control's mouse status</em>
	///
	/// \sa mouseStatus
	MouseStatus	mouseStatus_Calendar;
	/// \brief <em>Used to buffer the \c WM_LBUTTONDOWN message when applying the work-around for comctl32.dll 5.x</em>
	///
	/// \sa Flags.applyDropDownWorkaround, OnLButtonDown
	MSG bufferedLButtonDown;

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
		/// \brief <em>If \c TRUE, the current OLE drag'n'drop operation affects the drop-down calendar control</em>
		UINT isOverCalendar : 1;
		/// \brief <em>If \c TRUE, the \c ID_DRAGDROPDOWN timer is ticking</em>
		UINT dropDownTimerIsRunning : 1;
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
			isOverCalendar = FALSE;
			dropDownTimerIsRunning = FALSE;
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
			isOverCalendar = FALSE;
			dropDownTimerIsRunning = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa CalendarOLEDragLeaveOrDrop, OLEDragEnter
		HRESULT CalendarOLEDragEnter(void)
		{
			isOverCalendar = TRUE;
			autoScrolling.Reset();
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa CalendarOLEDragEnter, OLEDragLeaveOrDrop
		void CalendarOLEDragLeaveOrDrop(void)
		{
			isOverCalendar = FALSE;
			autoScrolling.Reset();
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop, CalendarOLEDragEnter
		HRESULT OLEDragEnter(void)
		{
			isOverCalendar = FALSE;
			dropDownTimerIsRunning = FALSE;
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter, CalendarOLEDragLeaveOrDrop
		void OLEDragLeaveOrDrop(void)
		{
			//
		}
	} dragDropStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-scroll the drop-down calendar control window during drag'n'drop</em>
		static const UINT_PTR ID_DRAGSCROLL = 13;
		/// \brief <em>The ID of the timer that is used to open the drop-down calendar control automatically during drag'n'drop</em>
		static const UINT_PTR ID_DRAGDROPDOWN = 14;
		/// \brief <em>The ID of the timer that is used to set the custom date format after control window recreation</em>
		///
		/// \remarks If the \c ID_REDRAW timer already is set, the \c ID_SETCUSTOMDATEFORMAT timer isn't set in
		///          order to save resources. The \c ID_REDRAW timer will execute the tasks, that otherwise
		///          would be executed by the \c ID_SETCUSTOMDATEFORMAT timer.
		static const UINT_PTR ID_SETCUSTOMDATEFORMAT = 15;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
		/// \brief <em>The interval of the timer that is used to set the custom date format after control window recreation</em>
		static const UINT INT_SETCUSTOMDATEFORMAT = 10;
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
};     // DateTimePicker

OBJECT_ENTRY_AUTO(__uuidof(DateTimePicker), DateTimePicker)