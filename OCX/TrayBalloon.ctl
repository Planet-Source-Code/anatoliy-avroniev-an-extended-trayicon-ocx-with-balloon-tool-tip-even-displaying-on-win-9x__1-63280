VERSION 5.00
Begin VB.UserControl ToolTipOnDemand 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   915
   ScaleWidth      =   945
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TT"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "ToolTipOnDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================
' Name:     ToolTipOnDemand.ctl
' Author:   Anatoliy Avroniev (aavroniev.axenet.ru)
' Date:     15 May 2006
' ======================================================================================

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

Public Enum blIconType
    blNoIcon = 0
    blIconInfo = 1
    blIconWarning = 2
    blIconError = 3
End Enum

Public Enum blStyleEnum
    blStandard = 0
    blBalloon = 1
End Enum

Public Enum blMouseEvent
    blMouseMove = WM_MOUSEMOVE
    blLeftButtonDown = WM_LBUTTONDOWN
    blLeftButtonUp = WM_LBUTTONUP
    blLeftButtonDoubleClick = WM_LBUTTONDBLCLK
    blRightButtonDown = WM_RBUTTONDOWN
    blRightButtonUp = WM_RBUTTONUP
    blRightButtonDoubleClick = WM_RBUTTONDBLCLK
    blMiddleButtonDown = WM_MBUTTONDOWN
    blMiddleButtonUp = WM_MBUTTONUP
    blMiddleButtonDoubleClick = WM_MBUTTONDBLCLK
End Enum

Dim WithEvents oTrayBalloon As cToolTipOnDemand
Attribute oTrayBalloon.VB_VarHelpID = -1
Private Const m_lWidth As Long = 375
Private Const m_lHeight As Long = 375

Private Const con_lToolTipCollor = &H80000018

Private m_lX As Long, m_lY As Long
Private m_Style As blStyleEnum
Private m_IconType As blIconType
Private m_lParentHwnd As Long
Private m_lForeColor As Long
Private m_lBackColor As Long
Private m_sTitle As String
Private m_sPrompt As String
Private m_bCentered As Boolean
Private m_lTimeOut As Long

Public Event MouseEvents(MouseEvent As blMouseEvent)
Public Event BalloonDestroyed()
Public Event BalloonShowed()

'//////////////////////////////////////////////////////
Public Property Let Style(ByVal vData As blStyleEnum)
  m_Style = vData
  Call PropertyChanged("Style")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.Style = m_Style
  End If
End Property
Public Property Get Style() As blStyleEnum
  Style = m_Style
End Property
'//////////////////////////////////////////////////////
Public Property Let IconType(ByVal vData As blIconType)
  m_IconType = vData
  Call PropertyChanged("IconType")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.IconType = m_IconType
  End If
End Property
Public Property Get IconType() As blIconType
  IconType = m_IconType
End Property
'//////////////////////////////////////////////////////
Public Property Let x(ByVal vData As Long)
  m_lX = vData
  Call PropertyChanged("x")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.x = vData
  End If
End Property
Public Property Get x() As Long
  x = m_lX
End Property
'//////////////////////////////////////////////////////
Public Property Let y(ByVal vData As Long)
  m_lY = vData
  Call PropertyChanged("y")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.y = vData
  End If
End Property
Public Property Get y() As Long
  y = m_lY
End Property
'//////////////////////////////////////////////////////
Public Property Let ParentHwnd(ByVal vData As Long)
  m_lParentHwnd = vData
  Call PropertyChanged("ParentHwnd")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.ParentHwnd = m_lParentHwnd
  End If
End Property
Public Property Get ParentHwnd() As Long
  ParentHwnd = m_lParentHwnd
End Property
'//////////////////////////////////////////////////////
Public Property Let ForeColor(ByVal vData As Long)
  m_lForeColor = vData
  Call PropertyChanged("ForeColor")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.ForeColor = m_lForeColor
  End If
End Property
Public Property Get ForeColor() As Long
  ForeColor = m_lForeColor
End Property
'//////////////////////////////////////////////////////
Public Property Let Title(ByVal vData As String)
  m_sTitle = vData
  Call PropertyChanged("Title")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.Title = m_sTitle
  End If
End Property
Public Property Get Title() As String
  Title = m_sTitle
End Property
'//////////////////////////////////////////////////////
Public Property Let BackColor(ByVal vData As Long)
  m_lBackColor = vData
  Call PropertyChanged("BackColor")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.BackColor = m_lBackColor
  End If
End Property
Public Property Get BackColor() As Long
  BackColor = m_lBackColor
End Property
'//////////////////////////////////////////////////////
Public Property Let Prompt(ByVal vData As String)
  m_sPrompt = vData
  Call PropertyChanged("Prompt")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.Prompt = m_sPrompt
  End If
End Property

Public Property Get Prompt() As String
  Prompt = m_sPrompt
End Property
'//////////////////////////////////////////////////////
Public Property Let Centered(ByVal vData As Boolean)
  m_bCentered = vData
  Call PropertyChanged("Centered")
  If Not oTrayBalloon Is Nothing Then
    oTrayBalloon.Centered = m_bCentered
  End If
End Property

Public Property Get Centered() As Boolean
  Centered = m_bCentered
End Property
'//////////////////////////////////////////////////////
Public Property Let TimeOut(ByVal vData As Long)
  m_lTimeOut = vData
  Call PropertyChanged("TimeOut")
End Property

Public Property Get TimeOut() As Long
  TimeOut = m_lTimeOut
End Property

Public Sub Show(Optional enIconType As blIconType = -1, Optional sPrompt As String = "~eMpTyStRiNg", Optional sTitle As String = "~eMpTyStRiNg", Optional lTimeout As Long = -1)
  
  If enIconType >= 0 Then m_IconType = enIconType
  If sPrompt <> "~eMpTyStRiNg" Then m_sPrompt = sPrompt
  If sTitle <> "~eMpTyStRiNg" Then m_sTitle = sTitle
  If lTimeout >= 0 Then m_lTimeOut = lTimeout
  
  Set oTrayBalloon = New cToolTipOnDemand
  With oTrayBalloon
    .x = m_lX
    .y = m_lY
    .Style = m_Style
    .ParentHwnd = m_lParentHwnd
    .ForeColor = m_lForeColor
    .BackColor = m_lBackColor
    .Centered = m_bCentered
  End With
  
  Call oTrayBalloon.Show(m_IconType, m_sPrompt, m_sTitle, m_lTimeOut)
End Sub

Public Sub Destroy()
  If Not oTrayBalloon Is Nothing Then
    Call oTrayBalloon.Destroy
    Set oTrayBalloon = Nothing
  End If

End Sub

Private Sub oTrayBalloon_BalloonDestroyed()
  RaiseEvent BalloonDestroyed
End Sub

Private Sub oTrayBalloon_BalloonShowed()
  RaiseEvent BalloonShowed
End Sub

Private Sub oTrayBalloon_MouseEvents(ByVal MouseEvent As Long)
  RaiseEvent MouseEvents(MouseEvent)
End Sub

Private Sub UserControl_Initialize()
  m_Style = blBalloon
  m_lBackColor = con_lToolTipCollor
  m_lForeColor = vbBlack

End Sub

Private Sub UserControl_Resize()
  UserControl.Width = m_lWidth
  UserControl.Height = m_lHeight
  Label1.Move 0, 0, m_lWidth, m_lHeight
End Sub

Private Sub UserControl_Terminate()
  Call Me.Destroy

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", m_Style, 1)
    Call PropBag.WriteProperty("IconType", m_IconType, 0)
    Call PropBag.WriteProperty("x", m_lX, 0)
    Call PropBag.WriteProperty("y", m_lY, 0)
    Call PropBag.WriteProperty("ParentHwnd", m_lParentHwnd, 0)
    Call PropBag.WriteProperty("ForeColor", m_lForeColor, vbBlack)
    Call PropBag.WriteProperty("BackColor", m_lBackColor, con_lToolTipCollor)
    Call PropBag.WriteProperty("Title", m_sTitle, "")
    Call PropBag.WriteProperty("Prompt", m_sPrompt, "")
    Call PropBag.WriteProperty("Centered", m_bCentered, False)
    Call PropBag.WriteProperty("TimeOut", m_lTimeOut, 0)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Style = PropBag.ReadProperty("Style", 1)
  m_IconType = PropBag.ReadProperty("IconType", 0)
  m_lX = PropBag.ReadProperty("x", 0)
  m_lY = PropBag.ReadProperty("y", 0)
  m_lParentHwnd = PropBag.ReadProperty("ParentHwnd", 0)
  m_lForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  m_lBackColor = PropBag.ReadProperty("m_lBackColor", con_lToolTipCollor)
  m_sTitle = PropBag.ReadProperty("Title", "")
  m_sPrompt = PropBag.ReadProperty("Prompt", "")
  m_bCentered = PropBag.ReadProperty("Centered", False)
  m_lTimeOut = PropBag.ReadProperty("TimeOut", 0)
  
End Sub

