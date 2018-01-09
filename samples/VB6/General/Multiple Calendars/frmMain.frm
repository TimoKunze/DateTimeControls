VERSION 5.00
Object = "{9812DC25-F663-48D3-AE15-E454712E6158}#1.5#0"; "DTCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "DateTimeControls 1.5 - Multiple Calendars Sample"
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
   Begin DTCtlsLibUCtl.Calendar CalU 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9375
      _cx             =   16536
      _cy             =   9128
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      DetectDoubleClicks=   0   'False
      DisabledEvents  =   11
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
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ScrollRate      =   0
      ShowTodayDate   =   -1  'True
      ShowTrailingDates=   -1  'True
      ShowWeekNumbers =   0   'False
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
      Left            =   3593
      TabIndex        =   0
      Top             =   5400
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type


  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
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
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  CalU.About
End Sub

Private Sub Form_Initialize()
  Dim bComctl32Version610OrNewer As Boolean
  Dim DLLVerData As DLLVERSIONINFO

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version610OrNewer = (.dwMajor > 6) Or ((.dwMajor = 6) And (.dwMinor >= 10))
  End With

  If Not bComctl32Version610OrNewer Then
    MsgBox "This sample requires version 6.10 or newer of comctl32.dll.", VbMsgBoxStyle.vbCritical, "Error"
    End
  End If
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass
End Sub

Private Sub Form_Resize()
  Dim b As Long
  Dim l As Long
  Dim minimumMCHeight As Long
  Dim minimumMCWidth As Long
  Dim r As Long
  Dim t As Long

  If Me.WindowState <> vbMinimized Then
    CalU.GetMinimumSize minimumMCWidth, minimumMCHeight
    If Me.ScaleWidth < minimumMCWidth + 50 Then
      Me.Width = Me.ScaleX(minimumMCWidth + 50, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    End If
    If Me.ScaleHeight < minimumMCHeight + cmdAbout.Height + 70 Then
      Me.Height = Me.ScaleY(minimumMCHeight + cmdAbout.Height + 70, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
    End If

    r = Me.ScaleWidth - 16
    b = Me.ScaleHeight - 8 - cmdAbout.Height - 16
    CalU.MinimizeRectangle l, t, r, b
    CalU.Move (Me.ScaleWidth - (r - l)) / 2, 8, r - l, b - t
    cmdAbout.Move (Me.ScaleWidth - cmdAbout.Width) / 2, Me.ScaleHeight - 8 - cmdAbout.Height
  End If
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
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong CalU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
