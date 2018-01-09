//////////////////////////////////////////////////////////////////////
/// \class Proxy_IDateTimePickerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IDateTimePickerEvents interface</em>
///
/// \if UNICODE
///   \sa DateTimePicker, DTCtlsLibU::_IDateTimePickerEvents
/// \else
///   \sa DateTimePicker, DTCtlsLibA::_IDateTimePickerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IDateTimePickerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IDateTimePickerEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c CalendarClick event</em>
	///
	/// \param[in] date The clicked date. May be zero.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the drop-down calendar control that was clicked. Any of the
	///            values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarClick, DateTimePicker::Raise_CalendarClick,
	///       Fire_CalendarMClick, Fire_CalendarRClick, Fire_Click, Fire_CalendarXClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarClick, DateTimePicker::Raise_CalendarClick,
	///       Fire_CalendarMClick, Fire_CalendarRClick, Fire_Click, Fire_CalendarXClick
	/// \endif
	HRESULT Fire_CalendarClick(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarCloseUp event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarCloseUp, DateTimePicker::Raise_CalendarCloseUp
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarCloseUp, DateTimePicker::Raise_CalendarCloseUp
	/// \endif
	HRESULT Fire_CalendarCloseUp(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARCLOSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarContextMenu event</em>
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
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarContextMenu,
	///       DateTimePicker::Raise_CalendarContextMenu, Fire_CalendarRClick, Fire_ContextMenu
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarContextMenu,
	///       DateTimePicker::Raise_CalendarContextMenu, Fire_CalendarRClick, Fire_ContextMenu
	/// \endif
	HRESULT Fire_CalendarContextMenu(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6].date = date;																	p[6].vt = VT_DATE;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pShowDefaultMenu;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarDateMouseEnter event</em>
	///
	/// \param[in] date The date that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Most of the values defined by the \c HitTestConstants enumeration are
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarDateMouseEnter,
	///       DateTimePicker::Raise_CalendarDateMouseEnter, Fire_CalendarDateMouseLeave,
	///       Fire_CalendarMouseMove
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarDateMouseEnter,
	///       DateTimePicker::Raise_CalendarDateMouseEnter, Fire_CalendarDateMouseLeave,
	///       Fire_CalendarMouseMove
	/// \endif
	HRESULT Fire_CalendarDateMouseEnter(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARDATEMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarDateMouseLeave event</em>
	///
	/// \param[in] date The date that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Most of the values defined by the \c HitTestConstants enumeration are
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarDateMouseLeave,
	///       DateTimePicker::Raise_CalendarDateMouseLeave, Fire_CalendarDateMouseEnter,
	///       Fire_CalendarMouseMove
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarDateMouseLeave,
	///       DateTimePicker::Raise_CalendarDateMouseLeave, Fire_CalendarDateMouseEnter,
	///       Fire_CalendarMouseMove
	/// \endif
	HRESULT Fire_CalendarDateMouseLeave(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARDATEMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarDropDown event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarDropDown, DateTimePicker::Raise_CalendarDropDown
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarDropDown, DateTimePicker::Raise_CalendarDropDown
	/// \endif
	HRESULT Fire_CalendarDropDown(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARDROPDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarKeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarKeyDown, DateTimePicker::Raise_CalendarKeyDown,
	///       Fire_CalendarKeyUp, Fire_CalendarKeyPress, Fire_KeyDown
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarKeyDown, DateTimePicker::Raise_CalendarKeyDown,
	///       Fire_CalendarKeyUp, Fire_CalendarKeyPress, Fire_KeyDown
	/// \endif
	HRESULT Fire_CalendarKeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARKEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarKeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarKeyPress, DateTimePicker::Raise_CalendarKeyPress,
	///       Fire_CalendarKeyDown, Fire_CalendarKeyUp, Fire_KeyPress
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarKeyPress, DateTimePicker::Raise_CalendarKeyPress,
	///       Fire_CalendarKeyDown, Fire_CalendarKeyUp, Fire_KeyPress
	/// \endif
	HRESULT Fire_CalendarKeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARKEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarKeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarKeyUp, DateTimePicker::Raise_CalendarKeyUp,
	///       Fire_CalendarKeyDown, Fire_CalendarKeyPress, Fire_KeyUp
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarKeyUp, DateTimePicker::Raise_CalendarKeyUp,
	///       Fire_CalendarKeyDown, Fire_CalendarKeyPress, Fire_KeyUp
	/// \endif
	HRESULT Fire_CalendarKeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARKEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMClick event</em>
	///
	/// \param[in] date The clicked date. May be zero.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the drop-down calendar control that was clicked. Any of the
	///            values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMClick, DateTimePicker::Raise_CalendarMClick,
	///       Fire_CalendarClick, Fire_CalendarRClick, Fire_MClick, Fire_CalendarXClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMClick, DateTimePicker::Raise_CalendarMClick,
	///       Fire_CalendarClick, Fire_CalendarRClick, Fire_MClick, Fire_CalendarXClick
	/// \endif
	HRESULT Fire_CalendarMClick(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMouseDown event</em>
	///
	/// \param[in] date The date that the mouse cursor is located over. May be zero.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseDown, DateTimePicker::Raise_CalendarMouseDown,
	///       Fire_CalendarMouseUp, Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarRClick,
	///       Fire_CalendarXClick, Fire_MouseDown
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseDown, DateTimePicker::Raise_CalendarMouseDown,
	///       Fire_CalendarMouseUp, Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarRClick,
	///       Fire_CalendarXClick, Fire_MouseDown
	/// \endif
	HRESULT Fire_CalendarMouseDown(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMouseEnter event</em>
	///
	/// \param[in] date The date that the mouse cursor is located over. May be zero.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseEnter,
	///       DateTimePicker::Raise_CalendarMouseEnter, Fire_CalendarMouseLeave, Fire_CalendarDateMouseEnter,
	///       Fire_CalendarMouseHover, Fire_CalendarMouseMove, Fire_MouseEnter
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseEnter,
	///       DateTimePicker::Raise_CalendarMouseEnter, Fire_CalendarMouseLeave, Fire_CalendarDateMouseEnter,
	///       Fire_CalendarMouseHover, Fire_CalendarMouseMove, Fire_MouseEnter
	/// \endif
	HRESULT Fire_CalendarMouseEnter(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMouseHover event</em>
	///
	/// \param[in] date The date that the mouse cursor is located over. May be zero.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseHover,
	///       DateTimePicker::Raise_CalendarMouseHover, Fire_CalendarMouseEnter, Fire_CalendarMouseLeave,
	///       Fire_CalendarMouseMove, Fire_MouseHover
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseHover,
	///       DateTimePicker::Raise_CalendarMouseHover, Fire_CalendarMouseEnter, Fire_CalendarMouseLeave,
	///       Fire_CalendarMouseMove, Fire_MouseHover
	/// \endif
	HRESULT Fire_CalendarMouseHover(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMouseLeave event</em>
	///
	/// \param[in] date The date that the mouse cursor is located over. May be zero.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseLeave, DateTimePicker::Raise_MouseLeave,
	///       Fire_CalendarMouseEnter, Fire_CalendarDateMouseLeave, Fire_CalendarMouseHover,
	///       Fire_CalendarMouseMove, Fire_MouseLeave
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseLeave, DateTimePicker::Raise_MouseLeave,
	///       Fire_CalendarMouseEnter, Fire_CalendarDateMouseLeave, Fire_CalendarMouseHover,
	///       Fire_CalendarMouseMove, Fire_MouseLeave
	/// \endif
	HRESULT Fire_CalendarMouseLeave(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMouseMove event</em>
	///
	/// \param[in] date The date that the mouse cursor is located over. May be zero.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseMove, DateTimePicker::Raise_CalendarMouseMove,
	///       Fire_CalendarMouseEnter, Fire_CalendarMouseLeave, Fire_CalendarMouseDown, Fire_CalendarMouseUp,
	///       Fire_CalendarMouseWheel, Fire_MouseMove
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseMove, DateTimePicker::Raise_CalendarMouseMove,
	///       Fire_CalendarMouseEnter, Fire_CalendarMouseLeave, Fire_CalendarMouseDown, Fire_CalendarMouseUp,
	///       Fire_CalendarMouseWheel, Fire_MouseMove
	/// \endif
	HRESULT Fire_CalendarMouseMove(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMouseUp event</em>
	///
	/// \param[in] date The date that the mouse cursor is located over. May be zero.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseUp, DateTimePicker::Raise_CalendarMouseUp,
	///       Fire_CalendarMouseDown, Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarRClick,
	///       Fire_CalendarXClick, Fire_MouseUp
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseUp, DateTimePicker::Raise_CalendarMouseUp,
	///       Fire_CalendarMouseDown, Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarRClick,
	///       Fire_CalendarXClick, Fire_MouseUp
	/// \endif
	HRESULT Fire_CalendarMouseUp(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarMouseWheel event</em>
	///
	/// \param[in] date The date that the mouse cursor is located over. May be zero.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            list box control's upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	/// \param[in] hitTestDetails The exact part of the drop-down calendar control that the mouse cursor's
	///            position lies in. Any of the values defined by the \c HitTestConstants enumeration is
	///            valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarMouseWheel,
	///       DateTimePicker::Raise_CalendarMouseWheel, Fire_CalendarMouseMove, Fire_MouseWheel
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarMouseWheel,
	///       DateTimePicker::Raise_CalendarMouseWheel, Fire_CalendarMouseMove, Fire_MouseWheel
	/// \endif
	HRESULT Fire_CalendarMouseWheel(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7].date = date;																	p[7].vt = VT_DATE;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(scrollAxis);				p[2].vt = VT_I4;
				p[1] = wheelDelta;																p[1].vt = VT_I2;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARMOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarOLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in] dropTarget The date that is below the mouse cursor's position.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails Specifies the part of the drop-down calendar control that the mouse
	///            cursor's position lies in. Any of the values defined by the \c HitTestConstants
	///            enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragDrop,
	///       DateTimePicker::Raise_CalendarOLEDragDrop, Fire_CalendarOLEDragEnter,
	///       Fire_CalendarOLEDragMouseMove, Fire_CalendarOLEDragLeave, Fire_CalendarMouseUp,
	///       Fire_OLEDragDrop
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragDrop,
	///       DateTimePicker::Raise_CalendarOLEDragDrop, Fire_CalendarOLEDragEnter,
	///       Fire_CalendarOLEDragMouseMove, Fire_CalendarOLEDragLeave, Fire_CalendarMouseUp,
	///       Fire_OLEDragDrop
	/// \endif
	HRESULT Fire_CalendarOLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, DATE dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].plVal = reinterpret_cast<PLONG>(pEffect);		p[6].vt = VT_I4 | VT_BYREF;
				p[5].date = dropTarget;														p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDAROLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarOLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in] dropTarget The date that is below the mouse cursor's position.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails Specifies the part of the drop-down calendar control that the mouse
	///            cursor's position lies in. Any of the values defined by the \c HitTestConstants
	///            enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the drop-down calendar control to the left; if set to a value greater
	///                than 0, it should scroll the drop-down calendar control to the right. The higher/lower
	///                the value is, the faster the caller should scroll the drop-down calendar control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragEnter,
	///       DateTimePicker::Raise_CalendarOLEDragEnter, Fire_CalendarOLEDragMouseMove,
	///       Fire_CalendarOLEDragLeave, Fire_CalendarOLEDragDrop, Fire_CalendarMouseEnter, Fire_OLEDragEnter
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragEnter,
	///       DateTimePicker::Raise_CalendarOLEDragEnter, Fire_CalendarOLEDragMouseMove,
	///       Fire_CalendarOLEDragLeave, Fire_CalendarOLEDragDrop, Fire_CalendarMouseEnter, Fire_OLEDragEnter
	/// \endif
	HRESULT Fire_CalendarOLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, DATE dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);		p[7].vt = VT_I4 | VT_BYREF;
				p[6].date = dropTarget;														p[6].vt = VT_DATE;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].plVal = pAutoHScrollVelocity;								p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDAROLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarOLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] dropTarget The date that is below the mouse cursor's position.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails Specifies the part of the drop-down calendar control that the mouse
	///            cursor's position lies in. Any of the values defined by the \c HitTestConstants
	///            enumeration is valid.
	/// \param[in,out] pAutoCloseUp If set to \c VARIANT_TRUE, the caller should close the drop-down calendar
	///                control; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragLeave,
	///       DateTimePicker::Raise_CalendarOLEDragLeave, Fire_CalendarOLEDragEnter,
	///       Fire_CalendarOLEDragMouseMove, Fire_CalendarOLEDragDrop, Fire_CalendarMouseLeave,
	///       Fire_OLEDragLeave
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragLeave,
	///       DateTimePicker::Raise_CalendarOLEDragLeave, Fire_CalendarOLEDragEnter,
	///       Fire_CalendarOLEDragMouseMove, Fire_CalendarOLEDragDrop, Fire_CalendarMouseLeave,
	///       Fire_OLEDragLeave
	/// \endif
	HRESULT Fire_CalendarOLEDragLeave(IOLEDataObject* pData, DATE dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoCloseUp)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].date = dropTarget;														p[6].vt = VT_DATE;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pAutoCloseUp;											p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDAROLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarOLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in] dropTarget The date that is below the mouse cursor's position.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the drop-down
	///            calendar control's upper-left corner.
	/// \param[in] hitTestDetails Specifies the part of the drop-down calendar control that the mouse
	///            cursor's position lies in. Any of the values defined by the \c HitTestConstants
	///            enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the drop-down calendar control to the left; if set to a value greater
	///                than 0, it should scroll the drop-down calendar control to the right. The higher/lower
	///                the value is, the faster the caller should scroll the drop-down calendar control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarOLEDragMouseMove,
	///       DateTimePicker::Raise_CalendarOLEDragMouseMove, Fire_CalendarOLEDragEnter,
	///       Fire_CalendarOLEDragLeave, Fire_CalendarOLEDragDrop, Fire_CalendarMouseMove,
	///       Fire_OLEDragMouseMove
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarOLEDragMouseMove,
	///       DateTimePicker::Raise_CalendarOLEDragMouseMove, Fire_CalendarOLEDragEnter,
	///       Fire_CalendarOLEDragLeave, Fire_CalendarOLEDragDrop, Fire_CalendarMouseMove,
	///       Fire_OLEDragMouseMove
	/// \endif
	HRESULT Fire_CalendarOLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, DATE dropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);		p[7].vt = VT_I4 | VT_BYREF;
				p[6].date = dropTarget;														p[6].vt = VT_DATE;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].plVal = pAutoHScrollVelocity;								p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDAROLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarRClick event</em>
	///
	/// \param[in] date The clicked date. May be zero.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the drop-down calendar control that was clicked. Any of the
	///            values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarRClick, DateTimePicker::Raise_CalendarRClick,
	///       Fire_CalendarContextMenu, Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarXClick,
	///       Fire_RClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarRClick, DateTimePicker::Raise_CalendarRClick,
	///       Fire_CalendarContextMenu, Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarXClick,
	///       Fire_RClick
	/// \endif
	HRESULT Fire_CalendarRClick(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARRCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarViewChanged event</em>
	///
	/// \param[in] oldView The previous view mode.
	/// \param[in] newView The new view mode.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarViewChanged,
	///       DateTimePicker::Raise_CalendarViewChanged
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarViewChanged,
	///       DateTimePicker::Raise_CalendarViewChanged
	/// \endif
	HRESULT Fire_CalendarViewChanged(ViewConstants oldView, ViewConstants newView)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].lVal = static_cast<LONG>(oldView);		p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(newView);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARVIEWCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CalendarXClick event</em>
	///
	/// \param[in] date The clicked date. May be zero.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the drop-down calendar
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the drop-down calendar control that was clicked. Any of the
	///            values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CalendarXClick, DateTimePicker::Raise_CalendarXClick,
	///       Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarRClick, Fire_XClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CalendarXClick, DateTimePicker::Raise_CalendarXClick,
	///       Fire_CalendarClick, Fire_CalendarMClick, Fire_CalendarRClick, Fire_XClick
	/// \endif
	HRESULT Fire_CalendarXClick(DATE date, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].date = date;																	p[5].vt = VT_DATE;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALENDARXCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CallbackFieldKeyDown event</em>
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
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CallbackFieldKeyDown,
	///       DateTimePicker::Raise_CallbackFieldKeyDown, Fire_KeyDown, Fire_GetCallbackFieldText
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CallbackFieldKeyDown,
	///       DateTimePicker::Raise_CallbackFieldKeyDown, Fire_KeyDown, Fire_GetCallbackFieldText
	/// \endif
	HRESULT Fire_CallbackFieldKeyDown(BSTR callbackField, SHORT keyCode, SHORT shift, DATE* pCurrentDate)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = callbackField;
				p[2] = keyCode;								p[2].vt = VT_I2;
				p[1] = shift;									p[1].vt = VT_I2;
				p[0].pdate = pCurrentDate;		p[0].vt = VT_DATE | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CALLBACKFIELDKEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::Click, DateTimePicker::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick, Fire_CalendarClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::Click, DateTimePicker::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick, Fire_CalendarClick
	/// \endif
	HRESULT Fire_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
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
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::ContextMenu, DateTimePicker::Raise_ContextMenu,
	///       Fire_RClick, Fire_CalendarContextMenu
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::ContextMenu, DateTimePicker::Raise_ContextMenu,
	///       Fire_RClick, Fire_CalendarContextMenu
	/// \endif
	HRESULT Fire_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;											p[4].vt = VT_I2;
				p[3] = shift;												p[3].vt = VT_I2;
				p[2] = x;														p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;														p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pShowDefaultMenu;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedCalendarControlWindow event</em>
	///
	/// \param[in] hWndCalendar The drop-down calendar control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CreatedCalendarControlWindow,
	///       DateTimePicker::Raise_CreatedCalendarControlWindow,
	///       Fire_DestroyedCalendarControlWindow, Fire_CalendarDropDown
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CreatedCalendarControlWindow,
	///       DateTimePicker::Raise_CreatedCalendarControlWindow,
	///       Fire_DestroyedCalendarControlWindow, Fire_CalendarDropDown
	/// \endif
	HRESULT Fire_CreatedCalendarControlWindow(LONG hWndCalendar)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndCalendar;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CREATEDCALENDARCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CreatedEditControlWindow,
	///       DateTimePicker::Raise_CreatedEditControlWindow,
	///       Fire_DestroyedEditControlWindow
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CreatedEditControlWindow,
	///       DateTimePicker::Raise_CreatedEditControlWindow,
	///       Fire_DestroyedEditControlWindow
	/// \endif
	HRESULT Fire_CreatedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CREATEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CurrentDateChanged event</em>
	///
	/// \param[in] dateSelected Specifies whether a date or the &ldquo;No selection&rdquo; state has been
	///            selected. If \c VARIANT_TRUE, the date specified by \c selectedDate has been selected;
	///            otherwise the &ldquo;No selection&rdquo; state has been selected.
	/// \param[in] selectedDate The date that has been selected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::CurrentDateChanged,
	///       DateTimePicker::Raise_CurrentDateChanged
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::CurrentDateChanged,
	///       DateTimePicker::Raise_CurrentDateChanged
	/// \endif
	HRESULT Fire_CurrentDateChanged(VARIANT_BOOL dateSelected, DATE selectedDate)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = dateSelected;
				p[0].date = selectedDate;		p[0].vt = VT_DATE;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DTPE_CURRENTDATECHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::DblClick, DateTimePicker::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::DblClick, DateTimePicker::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedCalendarControlWindow event</em>
	///
	/// \param[in] hWndCalendar The drop-down calendar control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::DestroyedCalendarControlWindow,
	///       DateTimePicker::Raise_DestroyedCalendarControlWindow,
	///       Fire_CreatedCalendarControlWindow, Fire_CalendarCloseUp
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::DestroyedCalendarControlWindow,
	///       DateTimePicker::Raise_DestroyedCalendarControlWindow,
	///       Fire_CreatedCalendarControlWindow, Fire_CalendarCloseUp
	/// \endif
	HRESULT Fire_DestroyedCalendarControlWindow(LONG hWndCalendar)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndCalendar;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_DESTROYEDCALENDARCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::DestroyedControlWindow,
	///       DateTimePicker::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::DestroyedControlWindow,
	///       DateTimePicker::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \endif
	HRESULT Fire_DestroyedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::DestroyedEditControlWindow,
	///       DateTimePicker::Raise_DestroyedEditControlWindow,
	///       Fire_CreatedEditControlWindow
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::DestroyedEditControlWindow,
	///       DateTimePicker::Raise_DestroyedEditControlWindow,
	///       Fire_CreatedEditControlWindow
	/// \endif
	HRESULT Fire_DestroyedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_DESTROYEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GetBoldDates event</em>
	///
	/// \param[in] firstDate The first date for which to retrieve the state.
	/// \param[in] numberOfDates The number of dates for which to retrieve the states.
	/// \param[in,out] ppStates An array containing the state for each date. The client may replace this
	///                array.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::GetBoldDates, DateTimePicker::Raise_GetBoldDates
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::GetBoldDates, DateTimePicker::Raise_GetBoldDates
	/// \endif
	HRESULT Fire_GetBoldDates(DATE firstDate, LONG numberOfDates, LPSAFEARRAY* ppStates)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].date = firstDate;			p[2].vt = VT_DATE;
				p[1] = numberOfDates;
				p[0].pparray = ppStates;		p[0].vt = VT_ARRAY | VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_DTPE_GETBOLDDATES, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GetCallbackFieldText event</em>
	///
	/// \param[in] callbackField The callback field for which the text is required. The number of \c X
	///            identifies the field.
	/// \param[in] dateToFormat The date to be formatted.
	/// \param[in,out] pTextHeight Receives the text to display in the callback field.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::GetCallbackFieldText,
	///       DateTimePicker::Raise_GetCallbackFieldText, Fire_GetCallbackFieldTextSize,
	///       Fire_CallbackFieldKeyDown
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::GetCallbackFieldText,
	///       DateTimePicker::Raise_GetCallbackFieldText, Fire_GetCallbackFieldTextSize,
	///       Fire_CallbackFieldKeyDown
	/// \endif
	HRESULT Fire_GetCallbackFieldText(BSTR callbackField, DATE dateToFormat, BSTR* pTextToDisplay)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = callbackField;
				p[1].date = dateToFormat;					p[1].vt = VT_DATE;
				p[0].pbstrVal = pTextToDisplay;		p[0].vt = VT_BSTR | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_DTPE_GETCALLBACKFIELDTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GetCallbackFieldTextSize event</em>
	///
	/// \param[in] callbackField The callback field for which the size is required. The number of \c X
	///            identifies the field.
	/// \param[in,out] pTextWidth Receives the maximum width of the callback field's content.
	/// \param[in,out] pTextHeight Receives the maximum height of the callback field's content.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::GetCallbackFieldTextSize,
	///       DateTimePicker::Raise_GetCallbackFieldTextSize, Fire_GetCallbackFieldText
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::GetCallbackFieldTextSize,
	///       DateTimePicker::Raise_GetCallbackFieldTextSize, Fire_GetCallbackFieldText
	/// \endif
	HRESULT Fire_GetCallbackFieldTextSize(BSTR callbackField, OLE_XSIZE_PIXELS* pTextWidth, OLE_YSIZE_PIXELS* pTextHeight)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = callbackField;
				p[1].plVal = pTextWidth;		p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pTextHeight;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_DTPE_GETCALLBACKFIELDTEXTSIZE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::KeyDown, DateTimePicker::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress, Fire_CalendarKeyDown, Fire_CallbackFieldKeyDown
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::KeyDown, DateTimePicker::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress, Fire_CalendarKeyDown, Fire_CallbackFieldKeyDown
	/// \endif
	HRESULT Fire_KeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DTPE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::KeyPress, DateTimePicker::Raise_KeyPress, Fire_KeyDown,
	///       Fire_KeyUp, Fire_CalendarKeyPress
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::KeyPress, DateTimePicker::Raise_KeyPress, Fire_KeyDown,
	///       Fire_KeyUp, Fire_CalendarKeyPress
	/// \endif
	HRESULT Fire_KeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::KeyUp, DateTimePicker::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress, Fire_CalendarKeyUp
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::KeyUp, DateTimePicker::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress, Fire_CalendarKeyUp
	/// \endif
	HRESULT Fire_KeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_DTPE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MClick, DateTimePicker::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick, Fire_CalendarMClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MClick, DateTimePicker::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick, Fire_CalendarMClick
	/// \endif
	HRESULT Fire_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MDblClick, DateTimePicker::Raise_MDblClick, Fire_MClick,
	///       Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MDblClick, DateTimePicker::Raise_MDblClick, Fire_MClick,
	///       Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MouseDown, DateTimePicker::Raise_MouseDown, Fire_MouseUp,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick, Fire_CalendarMouseDown
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MouseDown, DateTimePicker::Raise_MouseDown, Fire_MouseUp,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick, Fire_CalendarMouseDown
	/// \endif
	HRESULT Fire_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MouseEnter, DateTimePicker::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseHover, Fire_MouseMove, Fire_CalendarMouseEnter
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MouseEnter, DateTimePicker::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseHover, Fire_MouseMove, Fire_CalendarMouseEnter
	/// \endif
	HRESULT Fire_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MouseHover, DateTimePicker::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove, Fire_CalendarMouseHover
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MouseHover, DateTimePicker::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove, Fire_CalendarMouseHover
	/// \endif
	HRESULT Fire_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MouseLeave, DateTimePicker::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_MouseHover, Fire_MouseMove, Fire_CalendarMouseLeave
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MouseLeave, DateTimePicker::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_MouseHover, Fire_MouseMove, Fire_CalendarMouseLeave
	/// \endif
	HRESULT Fire_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MouseMove, DateTimePicker::Raise_MouseMove,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseDown, Fire_MouseWheel, Fire_MouseUp,
	///       Fire_CalendarMouseMove
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MouseMove, DateTimePicker::Raise_MouseMove,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseDown, Fire_MouseWheel, Fire_MouseUp,
	///       Fire_CalendarMouseMove
	/// \endif
	HRESULT Fire_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MouseUp, DateTimePicker::Raise_MouseUp, Fire_MouseDown,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick, Fire_CalendarMouseUp
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MouseUp, DateTimePicker::Raise_MouseUp, Fire_MouseDown,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick, Fire_CalendarMouseUp
	/// \endif
	HRESULT Fire_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseWheel event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::MouseWheel, DateTimePicker::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_CalendarMouseWheel
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::MouseWheel, DateTimePicker::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_CalendarMouseWheel
	/// \endif
	HRESULT Fire_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(scrollAxis);				p[1].vt = VT_I4;
				p[0] = wheelDelta;																p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_MOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::OLEDragDrop, DateTimePicker::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp,
	///       Fire_CalendarOLEDragDrop
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::OLEDragDrop, DateTimePicker::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp,
	///       Fire_CalendarOLEDragDrop
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pData;
				p[4].plVal = reinterpret_cast<PLONG>(pEffect);		p[4].vt = VT_I4 | VT_BYREF;
				p[3] = button;																		p[3].vt = VT_I2;
				p[2] = shift;																			p[2].vt = VT_I2;
				p[1] = x;																					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;																					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_DTPE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in,out] pAutoDropDown If set to \c VARIANT_TRUE, the caller should open the drop-down calendar
	///                control; otherwise it should cancel automatic drop-down.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::OLEDragEnter, DateTimePicker::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_CalendarOLEDragEnter
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::OLEDragEnter, DateTimePicker::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_CalendarOLEDragEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pAutoDropDown)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pData;
				p[5].plVal = reinterpret_cast<PLONG>(pEffect);		p[5].vt = VT_I4 | VT_BYREF;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pAutoDropDown;										p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_DTPE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::OLEDragLeave, DateTimePicker::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave,
	///       Fire_CalendarOLEDragLeave
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::OLEDragLeave, DateTimePicker::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave,
	///       Fire_CalendarOLEDragLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pData;
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_DTPE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in,out] pAutoDropDown If set to \c VARIANT_TRUE, the caller should open the drop-down calendar
	///                control; otherwise it should cancel automatic drop-down.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::OLEDragMouseMove, DateTimePicker::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove,
	///       Fire_CalendarOLEDragMouseMove
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::OLEDragMouseMove, DateTimePicker::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove,
	///       Fire_CalendarOLEDragMouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pAutoDropDown)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pData;
				p[5].plVal = reinterpret_cast<PLONG>(pEffect);		p[5].vt = VT_I4 | VT_BYREF;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pAutoDropDown;										p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_DTPE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ParseUserInput event</em>
	///
	/// \param[in] userInput The string that the user entered and that needs to be parsed.
	/// \param[out] pInputDate Must be set to the date that has been extracted from the user input. To
	///             indicate the &ldquo;No Selection&rdquo; state, this parameter can be set to 0.
	/// \param[out] pDateIsValid Specifies whether the entered string is a valid date. If set to
	///             \c VARIANT_TRUE, it is a valid date; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::ParseUserInput, DateTimePicker::Raise_ParseUserInput
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::ParseUserInput, DateTimePicker::Raise_ParseUserInput
	/// \endif
	HRESULT Fire_ParseUserInput(BSTR userInput, DATE* pInputDate, VARIANT_BOOL* pDateIsValid)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = userInput;
				p[1].pdate = pInputDate;				p[1].vt = VT_DATE | VT_BYREF;
				p[0].pboolVal = pDateIsValid;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_DTPE_PARSEUSERINPUT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::RClick, DateTimePicker::Raise_RClick, Fire_ContextMenu,
	///       Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick, Fire_CalendarRClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::RClick, DateTimePicker::Raise_RClick, Fire_ContextMenu,
	///       Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick, Fire_CalendarRClick
	/// \endif
	HRESULT Fire_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::RDblClick, DateTimePicker::Raise_RDblClick, Fire_RClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::RDblClick, DateTimePicker::Raise_RDblClick, Fire_RClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::RecreatedControlWindow,
	///       DateTimePicker::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::RecreatedControlWindow,
	///       DateTimePicker::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \endif
	HRESULT Fire_RecreatedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_DTPE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::ResizedControlWindow,
	///       DateTimePicker::Raise_ResizedControlWindow
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::ResizedControlWindow,
	///       DateTimePicker::Raise_ResizedControlWindow
	/// \endif
	HRESULT Fire_ResizedControlWindow(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_DTPE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be a
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::XClick, DateTimePicker::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_CalendarXClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::XClick, DateTimePicker::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick, Fire_CalendarXClick
	/// \endif
	HRESULT Fire_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
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
	/// \if UNICODE
	///   \sa DTCtlsLibU::_IDateTimePickerEvents::XDblClick, DateTimePicker::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa DTCtlsLibA::_IDateTimePickerEvents::XDblClick, DateTimePicker::Raise_XDblClick, Fire_XClick,
	///       Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_DTPE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IDateTimePickerEvents