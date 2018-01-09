VERSION 5.00
Object = "{9812DC25-F663-48D3-AE15-E454712E6158}#1.5#0"; "DTCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "DateTimeControls 1.5 - Custom Formats Sample"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1688
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin DTCtlsLibUCtl.DateTimePicker PickerU 
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      _cx             =   7435
      _cy             =   529
      AllowNullSelection=   0   'False
      Appearance      =   3
      BorderStyle     =   0
      CalendarBackColor=   -2147483643
      CalendarDragScrollTimeBase=   -1
      BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483640
      CalendarHighlightTodayDate=   -1  'True
      CalendarHoverTime=   -1
      CalendarKeepSelectionOnNavigation=   0   'False
      CalendarMonthBackColor=   -2147483643
      CalendarShowTodayDate=   -1  'True
      CalendarShowTrailingDates=   -1  'True
      CalendarShowWeekNumbers=   0   'False
      CalendarTitleBackColor=   -2147483646
      CalendarTitleForeColor=   -2147483639
      CalendarTrailingForeColor=   -2147483631
      CalendarUseShortestDayNames=   0   'False
      CalendarUseSystemFont=   -1  'True
      CalendarView    =   0
      DateFormat      =   0
      DateSelected    =   -1  'True
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   1087
      DontRedraw      =   0   'False
      DragDropDownTime=   -1
      DropDownAlignment=   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverTime       =   -1
      IMEMode         =   -1
      MaxDate         =   2958465
      MinDate         =   -53797
      MousePointer    =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      StartOfWeek     =   -1
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CustomDateFormat=   "frmMain.frx":0000
   End
   Begin DTCtlsLibUCtl.DateTimePicker PickerU 
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   4215
      _cx             =   7435
      _cy             =   529
      AllowNullSelection=   0   'False
      Appearance      =   3
      BorderStyle     =   0
      CalendarBackColor=   -2147483643
      CalendarDragScrollTimeBase=   -1
      BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483640
      CalendarHighlightTodayDate=   -1  'True
      CalendarHoverTime=   -1
      CalendarKeepSelectionOnNavigation=   0   'False
      CalendarMonthBackColor=   -2147483643
      CalendarShowTodayDate=   -1  'True
      CalendarShowTrailingDates=   -1  'True
      CalendarShowWeekNumbers=   0   'False
      CalendarTitleBackColor=   -2147483646
      CalendarTitleForeColor=   -2147483639
      CalendarTrailingForeColor=   -2147483631
      CalendarUseShortestDayNames=   0   'False
      CalendarUseSystemFont=   -1  'True
      CalendarView    =   0
      DateFormat      =   0
      DateSelected    =   -1  'True
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   1087
      DontRedraw      =   0   'False
      DragDropDownTime=   -1
      DropDownAlignment=   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverTime       =   -1
      IMEMode         =   -1
      MaxDate         =   2958465
      MinDate         =   -53797
      MousePointer    =   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      StartOfWeek     =   -1
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CustomDateFormat=   "frmMain.frx":0020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date and time:"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   645
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "With text:"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoW" (ByVal locale As Long, ByVal lcType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
  Private Declare Function GetUserDefaultLCID Lib "kernel32.dll" () As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const WM_WININICHANGE = &H1A
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)
      bCallDefProc = False
    Case WM_WININICHANGE
      ' SysDateTimePick32 wants this message to be forwarded
      SendMessageAsLong PickerU(0).hWnd, uMsg, wParam, lParam
      SendMessageAsLong PickerU(1).hWnd, uMsg, wParam, lParam
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  PickerU(0).About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass
  
  PickerU(0).CustomDateFormat = "'Today is: 'dddd MMM dd', 'yyyy"
  PickerU(0).DateFormat = dfCustom
  PickerU(1).CustomDateFormat = GetCurrentUserDateTimeFormat
  PickerU(1).DateFormat = dfCustom
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong PickerU(0).hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong PickerU(1).hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

Private Function GetCurrentUserDateTimeFormat() As String
  Const LOCALE_SSHORTDATE = &H1F
  Const LOCALE_SSHORTTIME = &H79
  Const LOCALE_STIMEFORMAT = &H1003
  Dim locale As Long
  Dim buffer As String
  Dim bufferSize As Long
  Dim ret As String
  
  locale = GetUserDefaultLCID
  
  bufferSize = GetLocaleInfo(locale, LOCALE_SSHORTDATE, 0, 0)
  If bufferSize = 0 Then
    locale = 1033     ' English (U.S.)
    bufferSize = GetLocaleInfo(locale, LOCALE_SSHORTDATE, 0, 0)
  End If
  buffer = String$(bufferSize + 1, Chr$(0))
  Call GetLocaleInfo(locale, LOCALE_SSHORTDATE, StrPtr(buffer), bufferSize)
  ret = Left$(buffer, bufferSize - 1)
  
  bufferSize = GetLocaleInfo(locale, LOCALE_SSHORTTIME, 0, 0)
  If bufferSize = 0 Then
    locale = 1033     ' English (U.S.)
    bufferSize = GetLocaleInfo(locale, LOCALE_SSHORTTIME, 0, 0)
  End If
  buffer = String$(bufferSize + 1, Chr$(0))
  Call GetLocaleInfo(locale, LOCALE_SSHORTTIME, StrPtr(buffer), bufferSize)
  ret = ret & "' '" & Left$(buffer, bufferSize - 1)
  
  GetCurrentUserDateTimeFormat = ret
End Function
