VERSION 5.00
Begin VB.UserControl TrayIcon 
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1005
   ScaleWidth      =   1455
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TI"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================
' Name:     TrayIcon.ctl
' Author:   Anatoliy Avroniev (aavroniev.axenet.ru)
' Date:     15 May 2006
' ======================================================================================

Public Type tiRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WM_USER As Long = &H400
Private Const NIN_SELECT = WM_USER
Private Const NIN_KEYSELECT = (WM_USER + 1)

Private Const WM_BALLOONSHOW = (WM_USER + 2)
Private Const WM_BALLOONHIDE = (WM_USER + 3)
Private Const WM_BALLOONRCLK = (WM_USER + 4)
Private Const WM_BALLOONLCLK = (WM_USER + 5)

Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3

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

Public Enum stBalloonIconType
    btNoIcon = NIIF_NONE
    btWarning = NIIF_WARNING
    btError = NIIF_ERROR
    btInfo = NIIF_INFO
End Enum

Public Enum stMouseEvent
    stMouseMove = WM_MOUSEMOVE
    stLeftButtonDown = WM_LBUTTONDOWN
    stLeftButtonUp = WM_LBUTTONUP
    stLeftButtonDoubleClick = WM_LBUTTONDBLCLK
    stRightButtonDown = WM_RBUTTONDOWN
    stRightButtonUp = WM_RBUTTONUP
    stRightButtonDoubleClick = WM_RBUTTONDBLCLK
    stMiddleButtonDown = WM_MBUTTONDOWN
    stMiddleButtonUp = WM_MBUTTONUP
    stMiddleButtonDoubleClick = WM_MBUTTONDBLCLK
End Enum

Public Enum stKeyEvent
    stSelect = NIN_SELECT
    stKeySelect = NIN_KEYSELECT
End Enum

Public Enum stBalloonClickType
    stbBalloonShow = WM_BALLOONSHOW
    stbBalloonHide = WM_BALLOONHIDE
    stbRightClick = WM_BALLOONRCLK
    stbLeftClick = WM_BALLOONLCLK
End Enum

Public Event TrayMouseEvent(ByVal MouseEvent As stMouseEvent)
Public Event TrayKeyEvent(ByVal KeyEvent As stKeyEvent)
Public Event TrayIconMoved(ByVal lX As Long, ByVal lY As Long)
Public Event BalloonClick(ByVal ClickType As stBalloonClickType)
Public Event TaskBarRecreated()

Dim WithEvents oTrayIcon As cTrayIcon
Attribute oTrayIcon.VB_VarHelpID = -1
Private m_bVisible As Boolean
Private m_sToolTipText As String
Private m_lIconHandle As Long
Private m_lIconMovementTrackInterval As Long
Private m_bTrackIconMovement As Boolean

Private Const m_lWidth As Long = 375
Private Const m_lHeight As Long = 375

'//////////////////////////////////////////////////////
Public Property Get IconMovementTrackInterval() As Long
  IconMovementTrackInterval = m_lIconMovementTrackInterval
End Property

Public Property Let IconMovementTrackInterval(lValue As Long)
  
  If lValue <= 0 Then
    m_lIconMovementTrackInterval = 1
  Else
    m_lIconMovementTrackInterval = lValue
  End If
  Call PropertyChanged("IconMovementTrackInterval")
  
  If Not oTrayIcon Is Nothing Then
    oTrayIcon.IconMovementTrackInterval = m_lIconMovementTrackInterval
  End If
  
End Property

'//////////////////////////////////////////////////////
Public Property Get TrackIconMovement() As Boolean
  TrackIconMovement = m_bTrackIconMovement
End Property

Public Property Let TrackIconMovement(bValue As Boolean)
  m_bTrackIconMovement = bValue
  Call PropertyChanged("TrackIconMovement")
  
  If Not oTrayIcon Is Nothing Then
    oTrayIcon.TrackIconMovement = m_bTrackIconMovement
  End If
End Property

'//////////////////////////////////////////////////////
Public Property Get Created() As Boolean
  If oTrayIcon Is Nothing Then
    Created = False
  Else
    Created = True
  End If
End Property

'//////////////////////////////////////////////////////
Public Property Let TrayIconVisible(vData As Boolean)
  m_bVisible = vData
  Call PropertyChanged("TrayIconVisible")
  If Not oTrayIcon Is Nothing Then
    oTrayIcon.Visible = vData
  End If
End Property
Public Property Get TrayIconVisible() As Boolean
  TrayIconVisible = m_bVisible
End Property

'//////////////////////////////////////////////////////
Public Property Let ToolTip(sToolTip As String)
  m_sToolTipText = sToolTip
  Call PropertyChanged("ToolTip")
  If Not oTrayIcon Is Nothing Then
    oTrayIcon.ToolTip = m_sToolTipText
  End If
End Property
Public Property Get ToolTip() As String
  ToolTip = m_sToolTipText
End Property

'//////////////////////////////////////////////////////
Property Let IconHandle(IconHandle As Long)
  m_lIconHandle = IconHandle
  If Not oTrayIcon Is Nothing Then
    oTrayIcon.IconHandle = m_lIconHandle
  End If
End Property
Property Get IconHandle() As Long
  IconHandle = m_lIconHandle
End Property

Public Function SysTrayHWnd() As Long
  If Not oTrayIcon Is Nothing Then
    SysTrayHWnd = oTrayIcon.SysTrayHWnd
  End If
End Function

Public Function CommonControlsVersion() As String
  CommonControlsVersion = FileVersion("COMCTL32.DLL")
End Function

Public Sub Create(lOwnerHwnd As Long, Optional hIcon As Long, Optional sToolTipText As String = "~nOtOoLtIp")
  If hIcon > 0 Then m_lIconHandle = hIcon
  If sToolTipText <> "~nOtOoLtIp" Then m_sToolTipText = sToolTipText
  
  Set oTrayIcon = New cTrayIcon
  With oTrayIcon
    .Visible = m_bVisible
    .ToolTip = m_sToolTipText
    .IconHandle = m_lIconHandle
    .TrackIconMovement = m_bTrackIconMovement
    .IconMovementTrackInterval = m_lIconMovementTrackInterval
    Call .Create(lOwnerHwnd, m_lIconHandle, m_sToolTipText)
  End With
End Sub

Public Sub Remove()
  If Not oTrayIcon Is Nothing Then
    Set oTrayIcon = Nothing
  End If
End Sub

Public Sub BalloonTipShow(Optional enIconType As blIconType = blNoIcon, Optional ByVal sPrompt As String, Optional ByVal sTitle As String, Optional TimeOut As Long)
  If Not oTrayIcon Is Nothing Then
    Call oTrayIcon.BalloonTipShow(enIconType, sPrompt, ByVal sTitle, TimeOut)
  End If
End Sub
    
Public Sub BalloonTipClose()
  If Not oTrayIcon Is Nothing Then
    Call oTrayIcon.BalloonTipClose
  End If
End Sub

Public Sub GetIconMiddle(lX As Long, lY As Long)
  If Not oTrayIcon Is Nothing Then
    Call oTrayIcon.GetIconMiddle(lX, lY)
  End If
End Sub

Public Function IconRect() As tiRECT
  Dim rct As tiRECT
  If Not oTrayIcon Is Nothing Then
    Call oTrayIcon.GetIconRect(rct.Left, rct.Right, rct.Top, rct.Bottom)
    IconRect = rct
  End If
End Function

Private Sub oTrayIcon_BalloonClick(ClickType As Long)
  RaiseEvent BalloonClick(ClickType)
End Sub

Private Sub oTrayIcon_TaskBarRecreated()
  RaiseEvent TaskBarRecreated
End Sub

Private Sub oTrayIcon_TrayIconMoved(lX As Long, lY As Long)
  RaiseEvent TrayIconMoved(lX, lY)
End Sub

Private Sub oTrayIcon_TrayKeyEvent(KeyEvent As Long)
  RaiseEvent TrayKeyEvent(KeyEvent)
End Sub

Private Sub oTrayIcon_TrayMouseEvent(MouseEvent As Long)
  RaiseEvent TrayMouseEvent(MouseEvent)
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = m_lWidth
  UserControl.Height = m_lHeight
  Label1.Move 0, 0, m_lWidth, m_lHeight
End Sub

Private Sub UserControl_Terminate()
  Call Me.Remove
End Sub

Private Sub UserControl_InitProperties()
  m_bVisible = True
  m_lIconMovementTrackInterval = 500
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TrayIconVisible", m_bVisible, True)
    Call PropBag.WriteProperty("ToolTip", m_sToolTipText, "")
    Call PropBag.WriteProperty("IconMovementTrackInterval", m_lIconMovementTrackInterval, 500)
    Call PropBag.WriteProperty("TrackIconMovement", m_bTrackIconMovement, False)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_bVisible = PropBag.ReadProperty("TrayIconVisible", True)
  m_sToolTipText = PropBag.ReadProperty("ToolTip", "")
  m_lIconMovementTrackInterval = PropBag.ReadProperty("IconMovementTrackInterval", 500)
  m_bTrackIconMovement = PropBag.ReadProperty("TrackIconMovement", False)
  
End Sub

