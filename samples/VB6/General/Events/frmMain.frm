VERSION 5.00
Object = "{9812DC25-F663-48D3-AE15-E454712E6158}#1.5#0"; "DTCtlsU.ocx"
Object = "{19AD6CAB-F3B8-4C05-81A1-55135F225D05}#1.5#0"; "DTCtlsA.ocx"
Begin VB.Form frmMain 
   Caption         =   "DateTimeControls 1.5 - Events Sample"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
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
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'Bildschirmmitte
   Begin DTCtlsLibACtl.Calendar CalA 
      Height          =   2820
      Left            =   2100
      TabIndex        =   3
      Top             =   3060
      Width           =   2850
      _cx             =   5027
      _cy             =   4974
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      DetectDoubleClicks=   0   'False
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
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
      ForeColor       =   -2147483640
      HighlightTodayDate=   -1  'True
      HoverTime       =   -1
      KeepSelectionOnNavigation=   0   'False
      MaxDate         =   2958465
      MaxSelectedDates=   7
      MinDate         =   -53797
      MonthBackColor  =   -2147483643
      MousePointer    =   0
      MultiSelect     =   0   'False
      Padding         =   4
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ScrollRate      =   0
      ShowTodayDate   =   -1  'True
      ShowTrailingDates=   -1  'True
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   -1
      SupportOLEDragImages=   -1  'True
      TitleBackColor  =   -2147483646
      TitleForeColor  =   -2147483639
      TrailingForeColor=   -2147483631
      UsePadding      =   0   'False
      UseShortestDayNames=   0   'False
      UseSystemFont   =   -1  'True
      View            =   0
   End
   Begin DTCtlsLibACtl.DateTimePicker PickerA 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   3060
      Width           =   1815
      _cx             =   3201
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
      DisabledEvents  =   1024
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
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      StartOfWeek     =   -1
      Style           =   0
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      CustomDateFormat=   "frmMain.frx":0000
   End
   Begin DTCtlsLibUCtl.Calendar CalU 
      Height          =   2820
      Left            =   2100
      TabIndex        =   1
      Top             =   120
      Width           =   2850
      _cx             =   5027
      _cy             =   4974
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      DetectDoubleClicks=   0   'False
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
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
      ForeColor       =   -2147483640
      HighlightTodayDate=   -1  'True
      HoverTime       =   -1
      KeepSelectionOnNavigation=   0   'False
      MaxDate         =   2958465
      MaxSelectedDates=   7
      MinDate         =   -53797
      MonthBackColor  =   -2147483643
      MousePointer    =   0
      MultiSelect     =   0   'False
      Padding         =   4
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ScrollRate      =   0
      ShowTodayDate   =   -1  'True
      ShowTrailingDates=   -1  'True
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   -1
      SupportOLEDragImages=   -1  'True
      TitleBackColor  =   -2147483646
      TitleForeColor  =   -2147483639
      TrailingForeColor=   -2147483631
      UsePadding      =   0   'False
      UseShortestDayNames=   0   'False
      UseSystemFont   =   -1  'True
      View            =   0
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   5400
      Width           =   975
   End
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
      Left            =   6150
      TabIndex        =   6
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin DTCtlsLibUCtl.DateTimePicker PickerU 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _cx             =   3201
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
      DisabledEvents  =   1024
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private bLog As Boolean
  Private objActiveCtl As Object


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
      SendMessageAsLong PickerA.hWnd, uMsg, wParam, lParam
      SendMessageAsLong PickerU.hWnd, uMsg, wParam, lParam
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked

  ' this is required to make the control work as expected
  Subclass
End Sub

Private Sub Form_Resize()
  Dim cx As Long
  Dim minimumMCHeight As Long
  Dim minimumMCWidth As Long

  If Me.WindowState <> vbMinimized Then
    CalU.GetMinimumSize minimumMCWidth, minimumMCHeight
    CalU.Move CalU.Left, CalU.Top, minimumMCWidth, minimumMCHeight
    CalA.Move CalU.Left, CalA.Top, minimumMCWidth, minimumMCHeight
    cx = Me.ScaleWidth - (CalU.Left + minimumMCWidth + 10)
    txtLog.Move Me.ScaleWidth - cx, 0, cx, Me.ScaleHeight - cmdAbout.Height - 10
    cmdAbout.Move txtLog.Left + (cx / 2) - cmdAbout.Width / 2, txtLog.Top + txtLog.Height + 5
    chkLog.Move txtLog.Left, cmdAbout.Top + 5
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub CalA_Click(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_Click: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_ContextMenu(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "CalA_ContextMenu: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub CalA_DateMouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_DateMouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_DateMouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_DateMouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_DblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_DblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CalA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub CalA_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "CalA_DragDrop"
End Sub

Private Sub CalA_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "CalA_DragOver"
End Sub

Private Sub CalA_GetBoldDates(ByVal firstDate As Date, ByVal numberOfDates As Long, states() As Boolean)
  AddLogEntry "CalA_GetBoldDates: firstDate=" & firstDate & ", numberOfDates=" & numberOfDates
End Sub

Private Sub CalA_GotFocus()
  AddLogEntry "CalA_GotFocus"
  Set objActiveCtl = CalA
End Sub

Private Sub CalA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CalA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CalA_KeyPress(keyAscii As Integer)
  AddLogEntry "CalA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub CalA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CalA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CalA_LostFocus()
  AddLogEntry "CalA_LostFocus"
End Sub

Private Sub CalA_MClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MDblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MDblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MouseDown(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MouseDown: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MouseHover(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MouseHover: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MouseMove(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
'  AddLogEntry "CalA_MouseMove: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MouseUp(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MouseUp: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_MouseWheel(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As DTCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_MouseWheel: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_OLEDragDrop(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CalA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    CalA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub CalA_OLEDragEnter(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CalA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub CalA_OLEDragLeave(ByVal data As DTCtlsLibACtl.IOLEDataObject, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CalA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub CalA_OLEDragMouseMove(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CalA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub CalA_RClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_RClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_RDblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_RDblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CalA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub CalA_ResizedControlWindow()
  AddLogEntry "CalA_ResizedControlWindow"
End Sub

Private Sub CalA_SelectionChanged(ByVal firstSelectedDate As Date, ByVal lastSelectedDate As Date)
  AddLogEntry "CalA_SelectionChanged: firstSelectedDate=" & firstSelectedDate & ", lastSelectedDate=" & lastSelectedDate
End Sub

Private Sub CalA_Validate(Cancel As Boolean)
  AddLogEntry "CalA_Validate"
End Sub

Private Sub CalA_ViewChanged(ByVal oldView As DTCtlsLibACtl.ViewConstants, ByVal newView As DTCtlsLibACtl.ViewConstants)
  AddLogEntry "CalA_ViewChanged: oldView=" & oldView & ", newView=" & newView
End Sub

Private Sub CalA_XClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_XClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalA_XDblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "CalA_XDblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_Click(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_Click: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_ContextMenu(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "CalU_ContextMenu: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub CalU_DateMouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_DateMouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_DateMouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_DateMouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_DblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_DblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CalU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub CalU_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "CalU_DragDrop"
End Sub

Private Sub CalU_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "CalU_DragOver"
End Sub

Private Sub CalU_GetBoldDates(ByVal firstDate As Date, ByVal numberOfDates As Long, states() As Boolean)
  AddLogEntry "CalU_GetBoldDates: firstDate=" & firstDate & ", numberOfDates=" & numberOfDates
End Sub

Private Sub CalU_GotFocus()
  AddLogEntry "CalU_GotFocus"
  Set objActiveCtl = CalU
End Sub

Private Sub CalU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CalU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CalU_KeyPress(keyAscii As Integer)
  AddLogEntry "CalU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub CalU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "CalU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub CalU_LostFocus()
  AddLogEntry "CalU_LostFocus"
End Sub

Private Sub CalU_MClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MDblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MDblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MouseDown(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MouseDown: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MouseHover(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MouseHover: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MouseMove(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
'  AddLogEntry "CalU_MouseMove: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MouseUp(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MouseUp: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_MouseWheel(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As DTCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_MouseWheel: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_OLEDragDrop(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CalU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    CalU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub CalU_OLEDragEnter(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CalU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub CalU_OLEDragLeave(ByVal data As DTCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "CalU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub CalU_OLEDragMouseMove(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "CalU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub CalU_RClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_RClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_RDblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_RDblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "CalU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub CalU_ResizedControlWindow()
  AddLogEntry "CalU_ResizedControlWindow"
End Sub

Private Sub CalU_SelectionChanged(ByVal firstSelectedDate As Date, ByVal lastSelectedDate As Date)
  AddLogEntry "CalU_SelectionChanged: firstSelectedDate=" & firstSelectedDate & ", lastSelectedDate=" & lastSelectedDate
End Sub

Private Sub CalU_Validate(Cancel As Boolean)
  AddLogEntry "CalU_Validate"
End Sub

Private Sub CalU_ViewChanged(ByVal oldView As DTCtlsLibUCtl.ViewConstants, ByVal newView As DTCtlsLibUCtl.ViewConstants)
  AddLogEntry "CalU_ViewChanged: oldView=" & oldView & ", newView=" & newView
End Sub

Private Sub CalU_XClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_XClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub CalU_XDblClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "CalU_XDblClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarCloseUp()
  AddLogEntry "PickerA_CalendarCloseUp"
End Sub

Private Sub PickerA_CalendarContextMenu(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "PickerA_CalendarContextMenu: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub PickerA_CalendarDateMouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarDateMouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarDateMouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarDateMouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarDropDown()
  AddLogEntry "PickerA_CalendarDropDown"
End Sub

Private Sub PickerA_CalendarKeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerA_CalendarKeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerA_CalendarKeyPress(keyAscii As Integer)
  AddLogEntry "PickerA_CalendarKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub PickerA_CalendarKeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerA_CalendarKeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerA_CalendarMClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarMClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarMouseDown(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarMouseDown: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarMouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarMouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarMouseHover(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarMouseHover: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarMouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarMouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarMouseMove(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
'  AddLogEntry "PickerA_CalendarMouseMove: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarMouseUp(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarMouseUp: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarMouseWheel(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As DTCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarMouseWheel: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarOLEDragDrop(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PickerA_CalendarOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    PickerA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub PickerA_CalendarOLEDragEnter(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "PickerA_CalendarOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub PickerA_CalendarOLEDragLeave(ByVal data As DTCtlsLibACtl.IOLEDataObject, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "PickerA_CalendarOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub PickerA_CalendarOLEDragMouseMove(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "PickerA_CalendarOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub PickerA_CalendarRClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarRClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CalendarViewChanged(ByVal oldView As DTCtlsLibACtl.ViewConstants, ByVal newView As DTCtlsLibACtl.ViewConstants)
  AddLogEntry "PickerA_CalendarViewChanged: oldView=" & oldView & ", newView=" & newView
End Sub

Private Sub PickerA_CalendarXClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibACtl.HitTestConstants)
  AddLogEntry "PickerA_CalendarXClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerA_CallbackFieldKeyDown(ByVal callbackField As String, ByVal keyCode As Integer, ByVal shift As Integer, CurrentDate As Date)
  AddLogEntry "PickerA_CallbackFieldKeyDown: callbackField=" & callbackField & ", keyCode=" & keyCode & ", shift=" & shift & ", currentDate=" & CurrentDate
End Sub

Private Sub PickerA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "PickerA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub PickerA_CreatedCalendarControlWindow(ByVal hWndCalendar As Long)
  AddLogEntry "PickerA_CreatedCalendarControlWindow: hWnd=0x" & Hex(hWndCalendar)
End Sub

Private Sub PickerA_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "PickerA_CreatedEditControlWindow: hWnd=0x" & Hex(hWndEdit)
End Sub

Private Sub PickerA_CurrentDateChanged(ByVal DateSelected As Boolean, ByVal selectedDate As Date)
  AddLogEntry "PickerA_CurrentDateChanged: dateSelected=" & DateSelected & ", selectedDate=" & selectedDate
End Sub

Private Sub PickerA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PickerA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PickerA_DestroyedCalendarControlWindow(ByVal hWndCalendar As Long)
  AddLogEntry "PickerA_DestroyedCalendarControlWindow: hWnd=0x" & Hex(hWndCalendar)
End Sub

Private Sub PickerA_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "PickerA_DestroyedEditControlWindow: hWnd=0x" & Hex(hWndEdit)
End Sub

Private Sub PickerA_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "PickerA_DragDrop"
End Sub

Private Sub PickerA_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "PickerA_DragOver"
End Sub

Private Sub PickerA_GetBoldDates(ByVal firstDate As Date, ByVal numberOfDates As Long, states() As Boolean)
  AddLogEntry "PickerA_GetBoldDates: firstDate=" & firstDate & ", numberOfDates=" & numberOfDates
End Sub

Private Sub PickerA_GetCallbackFieldText(ByVal callbackField As String, ByVal dateToFormat As Date, textToDisplay As String)
  AddLogEntry "PickerA_GetCallbackFieldText: callbackField=" & callbackField & ", dateToFormat=" & dateToFormat & ", textToDisplay=" & textToDisplay
End Sub

Private Sub PickerA_GetCallbackFieldTextSize(ByVal callbackField As String, textWidth As stdole.OLE_XSIZE_PIXELS, textHeight As stdole.OLE_YSIZE_PIXELS)
  AddLogEntry "PickerA_GetCallbackFieldTextSize: callbackField=" & callbackField & ", textWidth=" & textWidth & ", textHeight=" & textHeight
End Sub

Private Sub PickerA_GotFocus()
  AddLogEntry "PickerA_GotFocus"
  Set objActiveCtl = PickerA
End Sub

Private Sub PickerA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerA_KeyPress(keyAscii As Integer)
  AddLogEntry "PickerA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub PickerA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerA_LostFocus()
  AddLogEntry "PickerA_LostFocus"
End Sub

Private Sub PickerA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
'  AddLogEntry "PickerA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As DTCtlsLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "PickerA_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub PickerA_OLEDragDrop(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "PickerA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    PickerA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub PickerA_OLEDragEnter(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "PickerA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub PickerA_OLEDragLeave(ByVal data As DTCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "PickerA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub PickerA_OLEDragMouseMove(ByVal data As DTCtlsLibACtl.IOLEDataObject, effect As DTCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "PickerA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub PickerA_ParseUserInput(ByVal userInput As String, inputDate As Date, dateIsValid As Boolean)
  AddLogEntry "PickerA_ParseUserInput: userInput=" & userInput & ", inputDate=" & inputDate & ", dateIsValid=" & dateIsValid
  On Error Resume Next
  inputDate = DateValue(userInput)
  dateIsValid = (Err.Number = 0)
End Sub

Private Sub PickerA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PickerA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PickerA_ResizedControlWindow()
  AddLogEntry "PickerA_ResizedControlWindow"
End Sub

Private Sub PickerA_Validate(Cancel As Boolean)
  AddLogEntry "PickerA_Validate"
End Sub

Private Sub PickerA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_CalendarClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarCloseUp()
  AddLogEntry "PickerU_CalendarCloseUp"
End Sub

Private Sub PickerU_CalendarContextMenu(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  AddLogEntry "PickerU_CalendarContextMenu: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub PickerU_CalendarDateMouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarDateMouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarDateMouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarDateMouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarDropDown()
  AddLogEntry "PickerU_CalendarDropDown"
End Sub

Private Sub PickerU_CalendarKeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerU_CalendarKeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerU_CalendarKeyPress(keyAscii As Integer)
  AddLogEntry "PickerU_CalendarKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub PickerU_CalendarKeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerU_CalendarKeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerU_CalendarMClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarMClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarMouseDown(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarMouseDown: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarMouseEnter(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarMouseEnter: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarMouseHover(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarMouseHover: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarMouseLeave(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarMouseLeave: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarMouseMove(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
'  AddLogEntry "PickerU_CalendarMouseMove: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarMouseUp(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarMouseUp: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarMouseWheel(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As DTCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarMouseWheel: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarOLEDragDrop(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PickerU_CalendarOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    PickerU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub PickerU_CalendarOLEDragEnter(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "PickerU_CalendarOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub PickerU_CalendarOLEDragLeave(ByVal data As DTCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants, autoCloseUp As Boolean)
  Dim files() As String
  Dim str As String

  str = "PickerU_CalendarOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoCloseUp=" & autoCloseUp

  AddLogEntry str
End Sub

Private Sub PickerU_CalendarOLEDragMouseMove(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants, autoHScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "PickerU_CalendarOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget=" & dropTarget & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoHScrollVelocity=" & autoHScrollVelocity

  AddLogEntry str
End Sub

Private Sub PickerU_CalendarRClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarRClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CalendarViewChanged(ByVal oldView As DTCtlsLibUCtl.ViewConstants, ByVal newView As DTCtlsLibUCtl.ViewConstants)
  AddLogEntry "PickerU_CalendarViewChanged: oldView=" & oldView & ", newView=" & newView
End Sub

Private Sub PickerU_CalendarXClick(ByVal hitDate As Date, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As DTCtlsLibUCtl.HitTestConstants)
  AddLogEntry "PickerU_CalendarXClick: hitDate=" & hitDate & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
End Sub

Private Sub PickerU_CallbackFieldKeyDown(ByVal callbackField As String, ByVal keyCode As Integer, ByVal shift As Integer, CurrentDate As Date)
  AddLogEntry "PickerU_CallbackFieldKeyDown: callbackField=" & callbackField & ", keyCode=" & keyCode & ", shift=" & shift & ", currentDate=" & CurrentDate
End Sub

Private Sub PickerU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "PickerU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub PickerU_CreatedCalendarControlWindow(ByVal hWndCalendar As Long)
  AddLogEntry "PickerU_CreatedCalendarControlWindow: hWnd=0x" & Hex(hWndCalendar)
End Sub

Private Sub PickerU_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "PickerU_CreatedEditControlWindow: hWnd=0x" & Hex(hWndEdit)
End Sub

Private Sub PickerU_CurrentDateChanged(ByVal DateSelected As Boolean, ByVal selectedDate As Date)
  AddLogEntry "PickerU_CurrentDateChanged: dateSelected=" & DateSelected & ", selectedDate=" & selectedDate
End Sub

Private Sub PickerU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PickerU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PickerU_DestroyedCalendarControlWindow(ByVal hWndCalendar As Long)
  AddLogEntry "PickerU_DestroyedCalendarControlWindow: hWnd=0x" & Hex(hWndCalendar)
End Sub

Private Sub PickerU_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "PickerU_DestroyedEditControlWindow: hWnd=0x" & Hex(hWndEdit)
End Sub

Private Sub PickerU_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "PickerU_DragDrop"
End Sub

Private Sub PickerU_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "PickerU_DragOver"
End Sub

Private Sub PickerU_GetBoldDates(ByVal firstDate As Date, ByVal numberOfDates As Long, states() As Boolean)
  AddLogEntry "PickerU_GetBoldDates: firstDate=" & firstDate & ", numberOfDates=" & numberOfDates
End Sub

Private Sub PickerU_GetCallbackFieldText(ByVal callbackField As String, ByVal dateToFormat As Date, textToDisplay As String)
  AddLogEntry "PickerU_GetCallbackFieldText: callbackField=" & callbackField & ", dateToFormat=" & dateToFormat & ", textToDisplay=" & textToDisplay
End Sub

Private Sub PickerU_GetCallbackFieldTextSize(ByVal callbackField As String, textWidth As stdole.OLE_XSIZE_PIXELS, textHeight As stdole.OLE_YSIZE_PIXELS)
  AddLogEntry "PickerU_GetCallbackFieldTextSize: callbackField=" & callbackField & ", textWidth=" & textWidth & ", textHeight=" & textHeight
End Sub

Private Sub PickerU_GotFocus()
  AddLogEntry "PickerU_GotFocus"
  Set objActiveCtl = PickerU
End Sub

Private Sub PickerU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerU_KeyPress(keyAscii As Integer)
  AddLogEntry "PickerU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub PickerU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "PickerU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub PickerU_LostFocus()
  AddLogEntry "PickerU_LostFocus"
End Sub

Private Sub PickerU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
'  AddLogEntry "PickerU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_MouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As DTCtlsLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "PickerU_MouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub PickerU_OLEDragDrop(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "PickerU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    PickerU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub PickerU_OLEDragEnter(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "PickerU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub PickerU_OLEDragLeave(ByVal data As DTCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "PickerU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub PickerU_OLEDragMouseMove(ByVal data As DTCtlsLibUCtl.IOLEDataObject, effect As DTCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, autoDropDown As Boolean)
  Dim files() As String
  Dim str As String

  str = "PickerU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", autoDropDown=" & autoDropDown

  AddLogEntry str
End Sub

Private Sub PickerU_ParseUserInput(ByVal userInput As String, inputDate As Date, dateIsValid As Boolean)
  AddLogEntry "PickerU_ParseUserInput: userInput=" & userInput & ", inputDate=" & inputDate & ", dateIsValid=" & dateIsValid
  On Error Resume Next
  inputDate = DateValue(userInput)
  dateIsValid = (Err.Number = 0)
End Sub

Private Sub PickerU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PickerU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PickerU_ResizedControlWindow()
  AddLogEntry "PickerU_ResizedControlWindow"
End Sub

Private Sub PickerU_Validate(Cancel As Boolean)
  AddLogEntry "PickerU_Validate"
End Sub

Private Sub PickerU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub PickerU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "PickerU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
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
    SendMessageAsLong PickerU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong CalU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong PickerA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong CalA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
