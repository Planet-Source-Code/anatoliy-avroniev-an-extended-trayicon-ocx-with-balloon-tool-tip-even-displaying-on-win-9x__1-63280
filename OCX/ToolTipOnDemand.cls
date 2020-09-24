VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTipOnDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================
' Name:     cToolTipOnDemand class
' Author:   Anatoliy Avroniev (aavroniev.axenet.ru)
' Date:     15 May 2006
' ======================================================================================

'**********************************************************************************************************
'cTrayIcon class
'  Credits:
'    Nayan S. Patel: http://www.codeproject.com/Purgatory/Balloon-Tooltip.asp
'**********************************************************************************************************

' ***** MakeTopMost ******
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
             ByVal hWndInsertAfter As Long, _
             ByVal x As Long, ByVal y As Long, _
             ByVal cx As Long, ByVal cy As Long, _
             ByVal wFlags As Long) As Long

''Windows API Constants
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

'***************************
Private Const con_lToolTipCollor = &H80000018

'Initialization of New ClassNames
Private Const ICC_BAR_CLASSES = &H4      'toolbar, statusbar, trackbar, tooltips
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Type tagINITCOMMONCONTROLSEX
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be
                    ' initialized.
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
        (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Const WM_BALLOONLCLK = &H405
Private Const WM_BALLOONRCLK = &H404
Private Const WM_BALLOONXCLK = WM_BALLOONRCLK

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Private Const WM_TIMER = &H113

' Window Messages
Private Const WM_NULL = &H0
Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2
Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String, _
        ByVal dwStyle As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As Long, _
        ByVal hMenu As Long, _
        ByVal hInstance As Long, _
        ByRef lpParam As Any) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByRef lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Declare Function DestroyWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long

Private Declare Function ClientToScreen Lib "user32" ( _
        ByVal hwnd As Long, _
        lpPoint As POINTAPI) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

''Windows API Types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

'Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40 ' comctl32.dll v5.8 required
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTM_TRACKACTIVATE = WM_USER + 17
Private Const TTM_TRACKPOSITION = WM_USER + 18
Private Const WS_POPUP = &H80000000

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

'local variable(s) to hold property value(s)
Private mvarBackColor As Long
Private mvarTitle As String
Private mvarForeColor As Long
Private mvarIconType As balIconType
Private mvarCentered As Boolean
Private mvarPrompt As String
Private mvarStyle As balStyleEnum 'local copy
'Private mvarPopupOnDemand As Boolean

'private data
Private m_lTTHwnd As Long                                  ' hwnd of the tooltip
Private m_lParentHwnd As Long                              ' hwnd of the window the tooltip attached to
Private m_lX As Long, m_lY As Long
Private ti As TOOLINFO

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const IMAGE_BITMAP = 0&
Private Const IMAGE_ICON = 1&
Private Const IMAGE_CURSOR = 2&
Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_HINSTANCE = (-6)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


Private Declare Function SetTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long, _
   ByVal uElapse As Long, _
   ByVal lpTimerFunc As Long) As Long
   
Private Declare Function KillTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long) As Long

Private tmrRunning As Boolean, APP_TIMER_EVENT_ID As Long

Implements ISubclass

Public Event MouseEvents(ByVal MouseEvent As Long)
Public Event BalloonDestroyed()
Public Event BalloonShowed()

Private Sub InitComctl32(dwFlags As Long)
   On Error Resume Next
   Dim icc As tagINITCOMMONCONTROLSEX
   Dim lRetVal As Long
   
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   lRetVal = InitCommonControlsEx(icc)

   If lRetVal = 0 Then InitCommonControls
End Sub

Private Sub Class_Initialize()
  Call InitComctl32(ICC_BAR_CLASSES)
  mvarStyle = balBalloon
  Me.BackColor = con_lToolTipCollor
  Me.ForeColor = vbBlack
  'mvarCentered = True

End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  On Error GoTo lblErr:
   'Debug.Print "ISubclass_MsgResponse"
   
   Select Case CurrentMessage
   Case WM_TIMER
     ISubclass_MsgResponse = emrConsume
   Case WM_MOUSEMOVE To WM_MBUTTONDBLCLK
     ISubclass_MsgResponse = emrConsume
   Case Else
     ISubclass_MsgResponse = emrPreprocess
   End Select

lblExit:
  Exit Function
lblErr:
  Resume lblExit

End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
On Error GoTo lblErr:


If hwnd = m_lTTHwnd Then
  Select Case iMsg
    Case WM_TIMER
      Call Destroy
      Call ResetColors
      
    Case Else
      RaiseEvent MouseEvents(iMsg)

  End Select

Else
  ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
End If

lblExit:
  Exit Function
lblErr:
  Resume lblExit
  
End Function

Private Sub Class_Terminate()
    Call Destroy
End Sub

Public Property Let Style(ByVal vData As balStyleEnum)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Style = 5
    mvarStyle = vData
End Property

Public Property Get Style() As balStyleEnum
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Style
    Style = mvarStyle
End Property
'//////////////////////////////////////////////////////
Public Property Let IconType(ByVal vData As balIconType)
    mvarIconType = vData
    If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIconType <> balNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIconType), ByVal mvarTitle
    End If
End Property

Public Property Get IconType() As balIconType
    IconType = mvarIconType
End Property
'//////////////////////////////////////////////////////
Public Property Let x(ByVal vData As Long)
    m_lX = vData
    Call SetToolTipPosition(m_lX, m_lY)
End Property
Public Property Get x() As Long
    x = m_lX
End Property
'//////////////////////////////////////////////////////
Public Property Let y(ByVal vData As Long)
    m_lY = vData
    Call SetToolTipPosition(m_lX, m_lY)
End Property
Public Property Get y() As Long
    y = m_lY
End Property
'//////////////////////////////////////////////////////
Public Property Let ParentHwnd(ByVal vData As Long)
    m_lParentHwnd = vData
End Property
Public Property Get ParentHwnd() As Long
    ParentHwnd = m_lParentHwnd
End Property
'//////////////////////////////////////////////////////
Public Property Let ForeColor(ByVal vData As Long)
    mvarForeColor = TranslateColor(vData)
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If
End Property
Public Property Get ForeColor() As Long
    ForeColor = mvarForeColor
End Property
'//////////////////////////////////////////////////////
Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
    If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIconType <> balNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIconType), ByVal mvarTitle
    End If
End Property
Public Property Get Title() As String
    Title = ti.lpStr
End Property
'//////////////////////////////////////////////////////
Public Property Let BackColor(ByVal vData As Long)
    mvarBackColor = TranslateColor(vData)
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If
End Property
Public Property Get BackColor() As Long
    BackColor = mvarBackColor
End Property
'//////////////////////////////////////////////////////
Public Property Let Prompt(ByVal vData As String)
    mvarPrompt = vData
    ti.lpStr = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If
End Property

Public Property Get Prompt() As String
    Prompt = mvarPrompt
End Property
'//////////////////////////////////////////////////////
Public Property Let Centered(ByVal vData As Boolean)
    mvarCentered = vData
End Property

Public Property Get Centered() As Boolean
    Centered = mvarCentered
End Property

Public Sub Show(Optional ByVal enIconType As balIconType, Optional ByVal sPrompt As String, Optional ByVal sTitle As String, Optional ByVal TimeOut As Long)

  Dim lX As Long, lY As Long
    
  'If enIconType <> Empty Then Me.IconType = enIconType
  Me.IconType = enIconType
  If sPrompt <> Empty Then Me.Prompt = sPrompt
  If sTitle <> Empty Then Me.Title = sTitle
    
    'X and Y are in Pixel so dont send vbTwips value
    If CreateToolTip(m_lParentHwnd, m_lX, m_lY) Then
    
      AttachMessage Me, m_lTTHwnd, WM_MOUSEMOVE
      
      AttachMessage Me, m_lTTHwnd, WM_LBUTTONDOWN
      AttachMessage Me, m_lTTHwnd, WM_LBUTTONUP
      AttachMessage Me, m_lTTHwnd, WM_LBUTTONDBLCLK
      
      AttachMessage Me, m_lTTHwnd, WM_RBUTTONDOWN
      AttachMessage Me, m_lTTHwnd, WM_RBUTTONUP
      AttachMessage Me, m_lTTHwnd, WM_RBUTTONDBLCLK
      
      AttachMessage Me, m_lTTHwnd, WM_MBUTTONDOWN
      AttachMessage Me, m_lTTHwnd, WM_MBUTTONUP
      AttachMessage Me, m_lTTHwnd, WM_MBUTTONDBLCLK
      
      If TimeOut > 0 Then Call TimerBegin(TimeOut)
    
      RaiseEvent BalloonShowed
    End If
    
lblExit:

End Sub

Private Function CreateToolTip(ByVal lParentHwnd As Long, x As Long, y As Long) As Boolean
    Dim ptTip As Long
    Dim ret As Long
    Dim lWinStyle As Long
    Dim lpRect As RECT
    
    Call Destroy

    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    ''create baloon style if desired
    If mvarStyle = balBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON

    m_lTTHwnd = CreateWindowEx(0&, _
            TOOLTIPS_CLASSA, _
            vbNullString, _
            lWinStyle, _
            0&, _
            0&, _
            0&, _
            0&, _
            lParentHwnd, _
            0&, _
            App.hInstance, _
            0&)

    If m_lTTHwnd = 0 Then
      ret = mSystemUtils.GetLastError
      Err.Raise ret, App.Title, mSystemUtils.APIErrorDescription(ret)
      Exit Function
    End If
    
    ''now set our tooltip info structure
    With ti
        If mvarCentered Then
          .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_CENTERTIP 'Or TTF_ABSOLUTE
        Else
          .lFlags = TTF_IDISHWND Or TTF_TRACK 'Or TTF_ABSOLUTE
        End If

        'set the hwnd prop to our parent control's hwnd
        .hwnd = lParentHwnd
        .lId = lParentHwnd
        .hInstance = App.hInstance
        .lSize = Len(ti)
    End With

    If lParentHwnd <> 0 Then
      ''get the rect of the parent control
      GetClientRect lParentHwnd, lpRect
      ti.lpRect = lpRect
    End If
    
    ''add the tooltip structure
    ret = SendMessage(m_lTTHwnd, TTM_ADDTOOLA, 0&, ti)

    ''if we want a title or we want an icon
    If mvarTitle <> vbNullString Or mvarIconType <> balNoIcon Then
        ret = SendMessage(m_lTTHwnd, TTM_SETTITLE, CLng(mvarIconType), ByVal mvarTitle)
    End If

    If mvarForeColor <> Empty Then
        ret = SendMessage(m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&)
    End If

    If mvarBackColor <> Empty Then
        ret = SendMessage(m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&)
    End If
    
    Call SetToolTipPosition(x, y)

    If lParentHwnd = 0 Then
      'make tooltip window a topmost window
      'ret = SetWindowPos(m_lTTHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
      ret = SetWindowPos(m_lTTHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    
    CreateToolTip = True
End Function

Private Sub SetToolTipPosition(ByVal x As Long, ByVal y As Long)
    Dim ret As Long
    Dim ptTip As Long
    If m_lTTHwnd = 0 Then Exit Sub
    ptTip = y * &H10000 + x

    'These two messages will set the position of the tooltip:
    ret = SendMessage(m_lTTHwnd, TTM_TRACKPOSITION, 0, ByVal ptTip)
    ret = SendMessage(m_lTTHwnd, TTM_TRACKACTIVATE, True, ti)

End Sub

Public Sub Destroy()
    Call TimerStop
    
    If m_lTTHwnd <> 0 Then
      DetachMessage Me, m_lTTHwnd, WM_MOUSEMOVE
      
      DetachMessage Me, m_lTTHwnd, WM_LBUTTONDOWN
      DetachMessage Me, m_lTTHwnd, WM_LBUTTONUP
      DetachMessage Me, m_lTTHwnd, WM_LBUTTONDBLCLK
      
      DetachMessage Me, m_lTTHwnd, WM_RBUTTONDOWN
      DetachMessage Me, m_lTTHwnd, WM_RBUTTONUP
      DetachMessage Me, m_lTTHwnd, WM_RBUTTONDBLCLK
      
      DetachMessage Me, m_lTTHwnd, WM_MBUTTONDOWN
      DetachMessage Me, m_lTTHwnd, WM_MBUTTONUP
      DetachMessage Me, m_lTTHwnd, WM_MBUTTONDBLCLK
      
      DestroyWindow m_lTTHwnd
      m_lTTHwnd = 0
      RaiseEvent BalloonDestroyed
      
    End If
End Sub

Private Sub ResetColors()
    mvarForeColor = Empty
    mvarBackColor = Empty
End Sub
Private Sub TimerBegin(lMilliseconds As Long)
   
   If Not tmrRunning Then

      If lMilliseconds <> 0 Then
        
        APP_TIMER_EVENT_ID = 998
        tmrRunning = SetTimer(m_lTTHwnd, APP_TIMER_EVENT_ID, lMilliseconds, 0&)
        
        If tmrRunning Then
          AttachMessage Me, m_lTTHwnd, WM_TIMER
          'Debug.Print "timer started"
        End If
      End If
    End If
    
End Sub

Private Sub TimerStop()
   
   If tmrRunning Then
      DetachMessage Me, m_lTTHwnd, WM_TIMER
      Call KillTimer(m_lTTHwnd, APP_TIMER_EVENT_ID)
      'Debug.Print "timer stopped"
      tmrRunning = False
   End If

End Sub

