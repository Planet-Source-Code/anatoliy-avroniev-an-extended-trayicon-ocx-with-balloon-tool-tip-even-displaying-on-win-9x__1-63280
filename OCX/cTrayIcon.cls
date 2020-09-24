VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================
' Name:     cTrayIcon class
' Author:   Anatoliy Avroniev (aavroniev.axenet.ru)
' Date:     15 May 2006
' ======================================================================================

'**********************************************************************************************************
'cTrayIcon class
'  Credits:
'    Randy Birch: http://vbnet.mvps.org/index.html?code/subclass/shellnotifyballoontipclick.htm
'**********************************************************************************************************

Private Const WM_USER As Long = &H400
Private Const NIN_SELECT = WM_USER
Private Const NIN_KEYSELECT = (WM_USER + 1)

Private Const WM_BALLOONSHOW = (WM_USER + 2)
Private Const WM_BALLOONHIDE = (WM_USER + 3)
Private Const WM_BALLOONRCLK = (WM_USER + 4)
Private Const WM_BALLOONLCLK = (WM_USER + 5)


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
Private Const WM_CONTEXTMENU = &H7B

'general windows messages
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_COMMAND As Long = &H111
Private Const WM_CLOSE As Long = &H10
Private Const WM_TIMER = &H113
Private WM_TASKBARCREATED As Long

'private message the shell_notify api will pass
'to WindowProc when our systray icon is acted upon
Private Const WM_MYHOOK As Long = WM_USER + 1

Private Const APP_SYSTRAY_ID = 999

'flag preventing re-creating the timer
Private tmrRunning As Boolean
Private m_bVisible As Boolean
Private m_lIconHandle As Long
Private m_sToolTipText As String
Private m_lOwnerHwnd As Long

'ID constant representing this
'application for SetTimer
'Private Const APP_TIMER_EVENT_ID As Long = 998

Private ICON_MOVEMENT_TIMER_EVENT_ID As Long
Private bMovementTmrRunning As Boolean
Private m_lTrayIconX As Long, m_lTrayIconY As Long
Private m_lLastTrayIconX As Long, m_lLastTrayIconY As Long
Private m_bTrackIconMovement As Boolean
Private m_lIconMovementTrackInterval As Long

Private APP_TIMER_EVENT_ID As Long
'Private m_frmOwner As Form

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

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private Type NOTIFYICONDATA_9x
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Const NOTIFYICON_VERSION = 3
Private Const NOTIFYICON_OLDVERSION = 0
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3

Private Const NIIF_GUID = &H4

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private m_IconData As NOTIFYICONDATA


Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Private Declare Function SetTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long, _
   ByVal uElapse As Long, _
   ByVal lpTimerFunc As Long) As Long
   
Private Declare Function KillTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" Alias _
        "RegisterWindowMessageA" (ByVal lpString As String) As Long

Public Event TrayMouseEvent(MouseEvent As Long)
Public Event TrayKeyEvent(KeyEvent As Long)
Public Event TrayIconMoved(lX As Long, lY As Long)
Public Event BalloonClick(ClickType As Long)
Public Event TaskBarRecreated()

Implements ISubclass

Public Property Get IconMovementTrackInterval() As Long
  IconMovementTrackInterval = m_lIconMovementTrackInterval
End Property

Public Property Let IconMovementTrackInterval(lValue As Long)
  If lValue <= 0 Then
    m_lIconMovementTrackInterval = 1
  Else
    m_lIconMovementTrackInterval = lValue
  End If
  
  If m_IconData.hwnd > 0 Then
    If m_bTrackIconMovement Then
      Call IconMovementTimerStop
      Call IconMovementTimerBegin
    Else
      Call IconMovementTimerStop
    End If
  End If
End Property

Public Property Get TrackIconMovement() As Boolean
  TrackIconMovement = m_bTrackIconMovement
End Property

Public Property Let TrackIconMovement(bValue As Boolean)
  m_bTrackIconMovement = bValue
  
  If m_IconData.hwnd > 0 Then
    If m_bTrackIconMovement Then
      Call IconMovementTimerBegin
    Else
      Call IconMovementTimerStop
    End If
  End If
End Property

Public Property Get CommonControlsVersion() As String
  CommonControlsVersion = FileVersion("COMCTL32.DLL")
End Property

Private Function OSSupportsBalloons() As Boolean
  If tOSVERSIONINFO.lPlatformId = VER_PLATFORM_WIN32_NT Then
    If tOSVERSIONINFO.lMajorVersion >= 5 Then
      OSSupportsBalloons = True
    End If
  ElseIf tOSVERSIONINFO.lPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If tOSVERSIONINFO.lMajorVersion >= 4 And tOSVERSIONINFO.lMinorVersion = 90 Then
      OSSupportsBalloons = True
    End If
  End If
End Function

Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

Private Sub SetShellVersion()
  Dim nid9x As NOTIFYICONDATA_9x, lSize As Long
  lSize = Len(nid9x)
  
      If IsShellVersion(6) Then
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0 structure size
      
      ElseIf IsShellVersion(5) Then
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
      
      End If

End Sub

Private Sub Class_Initialize()
  
  'Debug.Print "Class_Initialize"
  m_bVisible = True
  Call mSystemUtils.GetOSVersion
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  On Error GoTo lblErr:
   
   Select Case CurrentMessage
   Case WM_TIMER Or WM_MYHOOK
     ISubclass_MsgResponse = emrConsume
   Case NIN_SELECT To WM_BALLOONLCLK
     ISubclass_MsgResponse = emrConsume
   Case WM_MOUSEMOVE To WM_MBUTTONDBLCLK
     ISubclass_MsgResponse = emrConsume
   Case WM_TASKBARCREATED
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

Select Case iMsg
  Case WM_TIMER
    'Debug.Print "WM_TIMER"
    If hwnd = m_lOwnerHwnd Then
      If wParam = APP_TIMER_EVENT_ID Then
        Call BalloonTipClose
      ElseIf wParam = ICON_MOVEMENT_TIMER_EVENT_ID Then
        'Debug.Print "ICON_MOVEMENT_TIMER_EVENT_ID"
        Call GetIconMiddle
        If (m_lLastTrayIconX <> m_lTrayIconX) Or (m_lLastTrayIconY <> m_lTrayIconY) Then
          m_lLastTrayIconX = m_lTrayIconX
          m_lLastTrayIconY = m_lTrayIconY
          RaiseEvent TrayIconMoved(m_lTrayIconX, m_lTrayIconY)
        Else
          m_lLastTrayIconX = m_lTrayIconX
          m_lLastTrayIconY = m_lTrayIconY
        End If
      
      Else
        ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
    Else
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
    End If
    
  Case WM_MYHOOK
    If hwnd = m_lOwnerHwnd Then
      If lParam >= WM_LBUTTONDOWN And lParam <= WM_MBUTTONDBLCLK Then
        Call BalloonTimerStop
        RaiseEvent TrayMouseEvent(lParam)
      
      ElseIf lParam >= NIN_SELECT And lParam <= NIN_KEYSELECT Then
        RaiseEvent TrayKeyEvent(lParam)
      
      ElseIf lParam >= WM_BALLOONSHOW And lParam <= WM_BALLOONHIDE Then
        RaiseEvent BalloonClick(lParam)
      
      ElseIf lParam >= WM_BALLOONRCLK And lParam <= WM_BALLOONLCLK Then
        Call BalloonTimerStop
        RaiseEvent BalloonClick(lParam)
      
      Else
        ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
    Else
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
    End If
    
  Case WM_TASKBARCREATED
    'Debug.Print "WM_TASKBARCREATED"
    Call ReCreate
    RaiseEvent TaskBarRecreated
  
  Case Else
    ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
End Select

lblExit:
  Exit Function
lblErr:
  Resume lblExit
  
End Function

Public Property Let Visible(vData As Boolean)
  If m_bVisible <> vData Then
    m_bVisible = vData
    If m_lOwnerHwnd = 0 Then Exit Property
    If vData Then
        m_IconData.hIcon = m_lIconHandle
    Else
        m_IconData.hIcon = 0
    End If
    m_IconData.uFlags = NIF_ICON Or NIF_TIP
    Shell_NotifyIcon NIM_MODIFY, m_IconData
  End If
  
End Property

Public Property Get Visible() As Boolean
  Visible = m_bVisible
End Property

Public Property Let ToolTip(Caption As String)
    
    With m_IconData
        .szTip = Caption & vbNullChar
        m_sToolTipText = Caption
        If m_lOwnerHwnd = 0 Then Exit Property
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData

End Property

Public Property Get ToolTip() As String
  ToolTip = m_IconData.szTip
End Property

Private Sub ReCreate()
  Call Create(m_lOwnerHwnd, , m_sToolTipText)
End Sub

Public Sub Create(ByVal lOwnerHwnd As Long, Optional ByVal hIcon As Long, Optional ByVal ToolTipText As String)

  Dim hIc As Long
  Dim sTTText As String
  m_lIconHandle = hIcon
  sTTText = ToolTipText
  
  If m_lOwnerHwnd > 0 Then Call Remove
  m_lOwnerHwnd = lOwnerHwnd
    
      Call SetShellVersion
    
      With m_IconData
        .cbSize = NOTIFYICONDATA_SIZE
        
        If m_bVisible Then .hIcon = m_lIconHandle Else .hIcon = 0
        
        If sTTText = "" Then
          .szTip = App.Title & vbNullChar
          m_sToolTipText = App.Title
        Else
          .szTip = sTTText & vbNullChar
          m_sToolTipText = sTTText
        End If
        
        .hwnd = m_lOwnerHwnd
        
        .uID = APP_SYSTRAY_ID
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallBackMessage = WM_MYHOOK
        
        .dwState = NIS_SHAREDICON
        .dwStateMask = 0
        .szInfo = Chr(0)
        .uTimeoutAndVersion = NOTIFYICON_VERSION
        .szInfoTitle = Chr(0)
        .dwInfoFlags = NIIF_NONE
      End With
      
      If m_lOwnerHwnd <> 0 Then
        Shell_NotifyIcon NIM_ADD, m_IconData
        AttachMessage Me, m_lOwnerHwnd, WM_MYHOOK
      End If
      
      If InitMessage Then
        AttachMessage Me, m_lOwnerHwnd, WM_TASKBARCREATED
      End If
        
      If m_bTrackIconMovement Then
        Call GetIconMiddle
        Call IconMovementTimerBegin
      End If
      
End Sub

Public Sub Remove()
    
    If m_lOwnerHwnd = 0 Then Exit Sub
    DetachMessage Me, m_lOwnerHwnd, WM_MYHOOK
    DetachMessage Me, m_lOwnerHwnd, WM_TASKBARCREATED
    Call IconMovementTimerStop
    Call BalloonTipClose
    
    Shell_NotifyIcon NIM_DELETE, m_IconData
    With m_IconData
        .hwnd = 0
        .uID = 0
        .uFlags = 0
        .uCallBackMessage = 0
        .hIcon = 0
        .szTip = Chr(0)
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Chr(0)
        .uTimeoutAndVersion = 0
        .szInfoTitle = Chr(0)
        .dwInfoFlags = 0
    End With
    m_lIconHandle = 0
    m_lOwnerHwnd = 0
    m_sToolTipText = ""
    
End Sub

Public Sub BalloonTipShow(Optional enIconType As balIconType = balNoIcon, Optional ByVal sPrompt As String, Optional ByVal sTitle As String, Optional TimeOut As Long)
  Dim sToolTip As String
    
  If m_lOwnerHwnd = 0 Then Exit Sub
  If sPrompt = Empty Then sPrompt = " "
  Call BalloonTipClose
  
      If OSSupportsBalloons Then
        With m_IconData
          .uFlags = NIF_INFO
          .szInfo = sPrompt & Chr(0)
          .szInfoTitle = sTitle & Chr(0)
          .dwInfoFlags = enIconType
          .uTimeoutAndVersion = NOTIFYICON_VERSION
        End With
    
        Shell_NotifyIcon NIM_MODIFY, m_IconData
        If TimeOut > 0 Then Call BalloonTimerBegin(TimeOut)
      
      End If
  
End Sub

Public Sub BalloonTipClose()
   
  If m_lOwnerHwnd = 0 Then Exit Sub
  Call BalloonTimerStop
  
  With m_IconData
    .uFlags = NIF_INFO
    .dwInfoFlags = 0
    .szInfoTitle = vbNullChar
    .szInfo = vbNullChar
  End With
  Call Shell_NotifyIcon(NIM_MODIFY, m_IconData)
  
End Sub

Property Get IconHandle() As Long
  IconHandle = m_lIconHandle
End Property

Property Let IconHandle(IconHandle As Long)
    
  m_lIconHandle = IconHandle
  If Not m_bVisible Then Exit Property
  With m_IconData
    .uFlags = NIF_ICON
    .hIcon = IconHandle
  End With
  If m_lOwnerHwnd = 0 Then Exit Property
  Shell_NotifyIcon NIM_MODIFY, m_IconData
End Property

Private Sub Class_Terminate()
  Call Remove
End Sub

Private Sub BalloonTimerBegin(lMilliseconds As Long)

   If m_IconData.hwnd = 0 Then Exit Sub
   If Not tmrRunning Then

      If lMilliseconds <> 0 Then
        
        APP_TIMER_EVENT_ID = 999
        tmrRunning = SetTimer(m_IconData.hwnd, APP_TIMER_EVENT_ID, lMilliseconds, 0&)
        
        If tmrRunning Then
          If Not bMovementTmrRunning Then
            AttachMessage Me, m_IconData.hwnd, WM_TIMER
          End If
          'Debug.Print "timer started"
        End If
      End If

   End If

End Sub

Private Sub BalloonTimerStop()
   
   If m_IconData.hwnd = 0 Then Exit Sub
   If tmrRunning = True Then
      If Not bMovementTmrRunning Then
        DetachMessage Me, m_IconData.hwnd, WM_TIMER
      End If
      Call KillTimer(m_IconData.hwnd, APP_TIMER_EVENT_ID)
      'Debug.Print "timer stopped"
      tmrRunning = False
   End If

End Sub

Public Sub GetIconMiddle(Optional lX As Long, Optional lY As Long)
      
    Dim rctTrayIcon As RECT, sToolTip As String
  
    sToolTip = Me.ToolTip
    Me.ToolTip = "wAnTeD tRaY iCoN " & App.hInstance
    rctTrayIcon = mTray.GetTrayIconRect
    Me.ToolTip = sToolTip
    
    With rctTrayIcon
      m_lTrayIconX = .Left + ((.Right - .Left) / 2)
      m_lTrayIconY = .Top + ((.Bottom - .Top) / 2)
      lX = m_lTrayIconX
      lY = m_lTrayIconY
    End With
   
End Sub

Public Sub GetIconRect(lLeft As Long, lRight As Long, lTop As Long, lBottom As Long)
  Dim rctTrayIcon As RECT, sToolTip As String
  
  sToolTip = Me.ToolTip
  Me.ToolTip = "wAnTeD tRaY iCoN " & App.hInstance
  rctTrayIcon = mTray.GetTrayIconRect
  Me.ToolTip = sToolTip
  
  With rctTrayIcon
    lLeft = .Left
    lRight = .Right
    lTop = .Bottom
    lBottom = .Bottom
  End With
End Sub

Public Property Get SysTrayHWnd() As Long
  SysTrayHWnd = mTray.GetSysTrayHWnd
End Property

'Register the windows message TaskbarCreated so we can watch for it
Private Function InitMessage() As Boolean
    
    If WM_TASKBARCREATED = 0 Then
      WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated")
      If WM_TASKBARCREATED <> 0 Then InitMessage = True
    Else
      InitMessage = True
    End If
    
End Function

Private Sub IconMovementTimerBegin()

   If m_IconData.hwnd = 0 Then Exit Sub
   If Not bMovementTmrRunning Then

        ICON_MOVEMENT_TIMER_EVENT_ID = 998
        bMovementTmrRunning = SetTimer(m_IconData.hwnd, ICON_MOVEMENT_TIMER_EVENT_ID, m_lIconMovementTrackInterval, 0&)
        
        If bMovementTmrRunning Then
          If Not tmrRunning Then
            AttachMessage Me, m_IconData.hwnd, WM_TIMER
          End If
          'Debug.Print "Movement timer started"
        End If

   End If

End Sub

Private Sub IconMovementTimerStop()
   
   If m_IconData.hwnd = 0 Then Exit Sub
   If bMovementTmrRunning = True Then
      If Not tmrRunning Then
        DetachMessage Me, m_IconData.hwnd, WM_TIMER
      End If
      Call KillTimer(m_IconData.hwnd, ICON_MOVEMENT_TIMER_EVENT_ID)
      'Debug.Print "Movement timer stopped"
      bMovementTmrRunning = False
   End If

End Sub


