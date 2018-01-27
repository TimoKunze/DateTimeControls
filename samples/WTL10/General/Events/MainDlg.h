// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////
#pragma once
#include <initguid.h>

#import <libid:9812DC25-F663-48d3-AE15-E454712E6158> version("1.5") raw_dispinterfaces
#import <libid:19AD6CAB-F3B8-4c05-81A1-55135F225D05> version("1.5") raw_dispinterfaces

DEFINE_GUID(LIBID_DTCtlsLibU, 0x9812DC25, 0xF663, 0x48d3, 0xAE, 0x15, 0xE4, 0x54, 0x71, 0x2E, 0x61, 0x58);
DEFINE_GUID(LIBID_DTCtlsLibA, 0x19AD6CAB, 0xF3B8, 0x4c05, 0x81, 0xA1, 0x55, 0x13, 0x5F, 0x22, 0x5D, 0x05);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_CALU, CMainDlg, &__uuidof(DTCtlsLibU::_ICalendarEvents), &LIBID_DTCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_PICKERU, CMainDlg, &__uuidof(DTCtlsLibU::_IDateTimePickerEvents), &LIBID_DTCtlsLibU, 1, 5>,
    public IDispEventImpl<IDC_CALA, CMainDlg, &__uuidof(DTCtlsLibA::_ICalendarEvents), &LIBID_DTCtlsLibA, 1, 5>,
    public IDispEventImpl<IDC_PICKERA, CMainDlg, &__uuidof(DTCtlsLibA::_IDateTimePickerEvents), &LIBID_DTCtlsLibA, 1, 5>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CButton> aboutButton;
	CContainedWindowT<CAxWindow> calUWnd;
	CContainedWindowT<CAxWindow> pickerUWnd;
	CContainedWindowT<CAxWindow> calAWnd;
	CContainedWindowT<CAxWindow> pickerAWnd;

	CMainDlg() :
	    aboutButton(this, 1),
	    calUWnd(this, 2),
	    pickerUWnd(this, 3),
	    calAWnd(this, 4),
	    pickerAWnd(this, 5)
	{
		hFocusedControl = NULL;
	}

	HWND hFocusedControl;

	struct Controls
	{
		CEdit logEdit;
		CComPtr<DTCtlsLibU::ICalendar> calU;
		CComPtr<DTCtlsLibU::IDateTimePicker> pickerU;
		CComPtr<DTCtlsLibA::ICalendar> calA;
		CComPtr<DTCtlsLibA::IDateTimePicker> pickerA;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusAbout)

		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		ALT_MSG_MAP(4)
		ALT_MSG_MAP(5)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 1, ClickCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 2, ContextMenuCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 3, DateMouseEnterCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 4, DateMouseLeaveCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 5, DblClickCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 6, DestroyedControlWindowCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 7, GetBoldDatesCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 8, KeyDownCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 9, KeyPressCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 10, KeyUpCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 11, MClickCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 12, MDblClickCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 13, MouseDownCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 14, MouseEnterCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 15, MouseHoverCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 16, MouseLeaveCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 17, MouseMoveCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 18, MouseUpCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 19, OLEDragDropCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 20, OLEDragEnterCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 21, OLEDragLeaveCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 22, OLEDragMouseMoveCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 23, RClickCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 24, RDblClickCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 25, RecreatedControlWindowCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 26, ResizedControlWindowCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 27, SelectionChangedCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 28, ViewChangedCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 29, MouseWheelCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 30, XClickCalU)
		SINK_ENTRY_EX(IDC_CALU, __uuidof(DTCtlsLibU::_ICalendarEvents), 31, XDblClickCalU)

		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 1, CalendarClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 2, CalendarCloseUpPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 3, CalendarContextMenuPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 4, CalendarDateMouseEnterPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 5, CalendarDateMouseLeavePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 7, CalendarDropDownPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 8, CalendarKeyDownPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 9, CalendarKeyPressPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 10, CalendarKeyUpPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 11, CalendarMClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 13, CalendarMouseDownPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 14, CalendarMouseEnterPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 15, CalendarMouseHoverPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 16, CalendarMouseLeavePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 17, CalendarMouseMovePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 18, CalendarMouseUpPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 19, CalendarOLEDragDropPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 20, CalendarOLEDragEnterPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 21, CalendarOLEDragLeavePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 22, CalendarOLEDragMouseMovePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 23, CalendarRClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 25, CalendarViewChangedPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 26, CallbackFieldKeyDownPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 27, ClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 28, ContextMenuPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 29, CreatedCalendarControlWindowPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 30, CreatedEditControlWindowPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 31, CurrentDateChangedPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 32, DblClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 33, DestroyedCalendarControlWindowPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 34, DestroyedControlWindowPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 35, DestroyedEditControlWindowPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 36, GetBoldDatesPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 37, GetCallbackFieldTextPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 38, GetCallbackFieldTextSizePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 39, KeyDownPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 40, KeyPressPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 41, KeyUpPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 42, MClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 43, MDblClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 44, MouseDownPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 45, MouseEnterPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 46, MouseHoverPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 47, MouseLeavePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 48, MouseMovePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 49, MouseUpPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 50, OLEDragDropPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 51, OLEDragEnterPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 52, OLEDragLeavePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 53, OLEDragMouseMovePickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 54, ParseUserInputPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 55, RClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 56, RDblClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 57, RecreatedControlWindowPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 58, ResizedControlWindowPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 59, CalendarMouseWheelPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 60, CalendarXClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 61, MouseWheelPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 62, XClickPickerU)
		SINK_ENTRY_EX(IDC_PICKERU, __uuidof(DTCtlsLibU::_IDateTimePickerEvents), 63, XDblClickPickerU)

		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 1, ClickCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 2, ContextMenuCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 3, DateMouseEnterCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 4, DateMouseLeaveCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 5, DblClickCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 6, DestroyedControlWindowCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 7, GetBoldDatesCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 8, KeyDownCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 9, KeyPressCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 10, KeyUpCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 11, MClickCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 12, MDblClickCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 13, MouseDownCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 14, MouseEnterCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 15, MouseHoverCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 16, MouseLeaveCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 17, MouseMoveCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 18, MouseUpCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 19, OLEDragDropCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 20, OLEDragEnterCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 21, OLEDragLeaveCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 22, OLEDragMouseMoveCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 23, RClickCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 24, RDblClickCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 25, RecreatedControlWindowCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 26, ResizedControlWindowCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 27, SelectionChangedCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 28, ViewChangedCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 29, MouseWheelCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 30, XClickCalA)
		SINK_ENTRY_EX(IDC_CALA, __uuidof(DTCtlsLibA::_ICalendarEvents), 31, XDblClickCalA)

		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 1, CalendarClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 2, CalendarCloseUpPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 3, CalendarContextMenuPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 4, CalendarDateMouseEnterPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 5, CalendarDateMouseLeavePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 7, CalendarDropDownPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 8, CalendarKeyDownPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 9, CalendarKeyPressPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 10, CalendarKeyUpPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 11, CalendarMClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 13, CalendarMouseDownPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 14, CalendarMouseEnterPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 15, CalendarMouseHoverPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 16, CalendarMouseLeavePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 17, CalendarMouseMovePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 18, CalendarMouseUpPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 19, CalendarOLEDragDropPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 20, CalendarOLEDragEnterPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 21, CalendarOLEDragLeavePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 22, CalendarOLEDragMouseMovePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 23, CalendarRClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 25, CalendarViewChangedPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 26, CallbackFieldKeyDownPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 27, ClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 28, ContextMenuPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 29, CreatedCalendarControlWindowPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 30, CreatedEditControlWindowPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 31, CurrentDateChangedPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 32, DblClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 33, DestroyedCalendarControlWindowPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 34, DestroyedControlWindowPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 35, DestroyedEditControlWindowPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 36, GetBoldDatesPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 37, GetCallbackFieldTextPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 38, GetCallbackFieldTextSizePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 39, KeyDownPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 40, KeyPressPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 41, KeyUpPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 42, MClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 43, MDblClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 44, MouseDownPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 45, MouseEnterPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 46, MouseHoverPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 47, MouseLeavePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 48, MouseMovePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 49, MouseUpPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 50, OLEDragDropPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 51, OLEDragEnterPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 52, OLEDragLeavePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 53, OLEDragMouseMovePickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 54, ParseUserInputPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 55, RClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 56, RDblClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 57, RecreatedControlWindowPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 58, ResizedControlWindowPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 59, CalendarMouseWheelPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 60, CalendarXClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 61, MouseWheelPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 62, XClickPickerA)
		SINK_ENTRY_EX(IDC_PICKERA, __uuidof(DTCtlsLibA::_IDateTimePickerEvents), 63, XDblClickPickerA)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusAbout(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);

	void __stdcall ClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DateMouseEnterCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DateMouseLeaveCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowCalU(long hWnd);
	void __stdcall GetBoldDatesCalU(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/);
	void __stdcall KeyDownCalU(short* keyCode, short shift);
	void __stdcall KeyPressCalU(short* keyAscii);
	void __stdcall KeyUpCalU(short* keyCode, short shift);
	void __stdcall MClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MDblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseDownCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseUpCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragDropCalU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterCalU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall OLEDragLeaveCalU(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragMouseMoveCalU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall RClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RDblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowCalU(long hWnd);
	void __stdcall ResizedControlWindowCalU();
	void __stdcall SelectionChangedCalU(DATE firstSelectedDate, DATE lastSelectedDate);
	void __stdcall ViewChangedCalU(DTCtlsLibU::ViewConstants oldView, DTCtlsLibU::ViewConstants newView);
	void __stdcall XClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall XDblClickCalU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);

	void __stdcall CalendarClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarCloseUpPickerU();
	void __stdcall CalendarContextMenuPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CalendarDateMouseEnterPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarDateMouseLeavePickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarDropDownPickerU();
	void __stdcall CalendarKeyDownPickerU(short* keyCode, short shift);
	void __stdcall CalendarKeyPressPickerU(short* keyAscii);
	void __stdcall CalendarKeyUpPickerU(short* keyCode, short shift);
	void __stdcall CalendarMClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseDownPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseEnterPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseHoverPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseLeavePickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseMovePickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseUpPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseWheelPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarOLEDragDropPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarOLEDragEnterPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall CalendarOLEDragLeavePickerU(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall CalendarOLEDragMouseMovePickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall CalendarRClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CalendarViewChangedPickerU(DTCtlsLibU::ViewConstants oldView, DTCtlsLibU::ViewConstants newView);
	void __stdcall CalendarXClickPickerU(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall CallbackFieldKeyDownPickerU(BSTR callbackField, short keyCode, short shift, DATE* currentDate);
	void __stdcall ClickPickerU(short button, short shift, long x, long y);
	void __stdcall ContextMenuPickerU(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedCalendarControlWindowPickerU(long hWndCalendar);
	void __stdcall CreatedEditControlWindowPickerU(long hWndEdit);
	void __stdcall CurrentDateChangedPickerU(VARIANT_BOOL dateSelected, DATE selectedDate);
	void __stdcall DblClickPickerU(short button, short shift, long x, long y);
	void __stdcall DestroyedCalendarControlWindowPickerU(long hWndCalendar);
	void __stdcall DestroyedControlWindowPickerU(long hWnd);
	void __stdcall DestroyedEditControlWindowPickerU(long hWndEdit);
	void __stdcall GetBoldDatesPickerU(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/);
	void __stdcall GetCallbackFieldTextPickerU(BSTR callbackField, DATE dateToFormat, BSTR* textToDisplay);
	void __stdcall GetCallbackFieldTextSizePickerU(BSTR callbackField, OLE_XSIZE_PIXELS* textWidth, OLE_YSIZE_PIXELS* textHeight);
	void __stdcall KeyDownPickerU(short* keyCode, short shift);
	void __stdcall KeyPressPickerU(short* keyAscii);
	void __stdcall KeyUpPickerU(short* keyCode, short shift);
	void __stdcall MClickPickerU(short button, short shift, long x, long y);
	void __stdcall MDblClickPickerU(short button, short shift, long x, long y);
	void __stdcall MouseDownPickerU(short button, short shift, long x, long y);
	void __stdcall MouseEnterPickerU(short button, short shift, long x, long y);
	void __stdcall MouseHoverPickerU(short button, short shift, long x, long y);
	void __stdcall MouseLeavePickerU(short button, short shift, long x, long y);
	void __stdcall MouseMovePickerU(short button, short shift, long x, long y);
	void __stdcall MouseUpPickerU(short button, short shift, long x, long y);
	void __stdcall MouseWheelPickerU(short button, short shift, long x, long y, DTCtlsLibU::ScrollAxisConstants scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterPickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragLeavePickerU(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMovePickerU(LPDISPATCH data, DTCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown);
	void __stdcall ParseUserInputPickerU(BSTR userInput, DATE* inputDate, VARIANT_BOOL* dateIsValid);
	void __stdcall RClickPickerU(short button, short shift, long x, long y);
	void __stdcall RDblClickPickerU(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowPickerU(long hWnd);
	void __stdcall ResizedControlWindowPickerU();
	void __stdcall XClickPickerU(short button, short shift, long x, long y);
	void __stdcall XDblClickPickerU(short button, short shift, long x, long y);

	void __stdcall ClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DateMouseEnterCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DateMouseLeaveCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowCalA(long hWnd);
	void __stdcall GetBoldDatesCalA(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/);
	void __stdcall KeyDownCalA(short* keyCode, short shift);
	void __stdcall KeyPressCalA(short* keyAscii);
	void __stdcall KeyUpCalA(short* keyCode, short shift);
	void __stdcall MClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MDblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseDownCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseUpCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragDropCalA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterCalA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall OLEDragLeaveCalA(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragMouseMoveCalA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall RClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RDblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowCalA(long hWnd);
	void __stdcall ResizedControlWindowCalA();
	void __stdcall SelectionChangedCalA(DATE firstSelectedDate, DATE lastSelectedDate);
	void __stdcall ViewChangedCalA(DTCtlsLibA::ViewConstants oldView, DTCtlsLibA::ViewConstants newView);
	void __stdcall XClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall XDblClickCalA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);

	void __stdcall CalendarClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarCloseUpPickerA();
	void __stdcall CalendarContextMenuPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CalendarDateMouseEnterPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarDateMouseLeavePickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarDropDownPickerA();
	void __stdcall CalendarKeyDownPickerA(short* keyCode, short shift);
	void __stdcall CalendarKeyPressPickerA(short* keyAscii);
	void __stdcall CalendarKeyUpPickerA(short* keyCode, short shift);
	void __stdcall CalendarMClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseDownPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseEnterPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseHoverPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseLeavePickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseMovePickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseUpPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarMouseWheelPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::ScrollAxisConstants scrollAxis, short wheelDelta, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarOLEDragDropPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarOLEDragEnterPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall CalendarOLEDragLeavePickerA(LPDISPATCH data, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* autoCloseUp);
	void __stdcall CalendarOLEDragMouseMovePickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, DATE dropTarget, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails, long* autoHScrollVelocity);
	void __stdcall CalendarRClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CalendarViewChangedPickerA(DTCtlsLibA::ViewConstants oldView, DTCtlsLibA::ViewConstants newView);
	void __stdcall CalendarXClickPickerA(DATE hitDate, short button, short shift, long x, long y, DTCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall CallbackFieldKeyDownPickerA(BSTR callbackField, short keyCode, short shift, DATE* currentDate);
	void __stdcall ClickPickerA(short button, short shift, long x, long y);
	void __stdcall ContextMenuPickerA(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedCalendarControlWindowPickerA(long hWndCalendar);
	void __stdcall CreatedEditControlWindowPickerA(long hWndEdit);
	void __stdcall CurrentDateChangedPickerA(VARIANT_BOOL dateSelected, DATE selectedDate);
	void __stdcall DblClickPickerA(short button, short shift, long x, long y);
	void __stdcall DestroyedCalendarControlWindowPickerA(long hWndCalendar);
	void __stdcall DestroyedControlWindowPickerA(long hWnd);
	void __stdcall DestroyedEditControlWindowPickerA(long hWndEdit);
	void __stdcall GetBoldDatesPickerA(DATE firstDate, long numberOfDates, LPSAFEARRAY* /*states*/);
	void __stdcall GetCallbackFieldTextPickerA(BSTR callbackField, DATE dateToFormat, BSTR* textToDisplay);
	void __stdcall GetCallbackFieldTextSizePickerA(BSTR callbackField, OLE_XSIZE_PIXELS* textWidth, OLE_YSIZE_PIXELS* textHeight);
	void __stdcall KeyDownPickerA(short* keyCode, short shift);
	void __stdcall KeyPressPickerA(short* keyAscii);
	void __stdcall KeyUpPickerA(short* keyCode, short shift);
	void __stdcall MClickPickerA(short button, short shift, long x, long y);
	void __stdcall MDblClickPickerA(short button, short shift, long x, long y);
	void __stdcall MouseDownPickerA(short button, short shift, long x, long y);
	void __stdcall MouseEnterPickerA(short button, short shift, long x, long y);
	void __stdcall MouseHoverPickerA(short button, short shift, long x, long y);
	void __stdcall MouseLeavePickerA(short button, short shift, long x, long y);
	void __stdcall MouseMovePickerA(short button, short shift, long x, long y);
	void __stdcall MouseUpPickerA(short button, short shift, long x, long y);
	void __stdcall MouseWheelPickerA(short button, short shift, long x, long y, DTCtlsLibA::ScrollAxisConstants scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterPickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown);
	void __stdcall OLEDragLeavePickerA(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMovePickerA(LPDISPATCH data, DTCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, VARIANT_BOOL* autoDropDown);
	void __stdcall ParseUserInputPickerA(BSTR userInput, DATE* inputDate, VARIANT_BOOL* dateIsValid);
	void __stdcall RClickPickerA(short button, short shift, long x, long y);
	void __stdcall RDblClickPickerA(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowPickerA(long hWnd);
	void __stdcall ResizedControlWindowPickerA();
	void __stdcall XClickPickerA(short button, short shift, long x, long y);
	void __stdcall XDblClickPickerA(short button, short shift, long x, long y);
};
