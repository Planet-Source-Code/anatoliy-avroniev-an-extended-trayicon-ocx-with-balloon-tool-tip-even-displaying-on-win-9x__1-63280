VERSION 5.00
Object = "*\A..\..\..\MYPROJ~2\COMPON~1\ASTRAY~3\OCX\TrayIconOCX.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   7395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7395
   StartUpPosition =   1  'CenterOwner
   Begin TrayIconOCX.ToolTipOnDemand ToolTipOnDemand1 
      Left            =   4560
      Top             =   120
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin TrayIconOCX.TrayIcon TrayIcon1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Timer tmrExit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6720
      Top             =   60
   End
   Begin VB.TextBox txtEvents 
      Height          =   1935
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   30
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Frame frTrayBalloon 
      Caption         =   "Tray Balloon"
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   7215
      Begin VB.TextBox txtTimeout 
         Height          =   285
         Left            =   2040
         TabIndex        =   32
         Text            =   "5000"
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   4800
         ScaleHeight     =   1575
         ScaleWidth      =   2295
         TabIndex        =   22
         Top             =   240
         Width           =   2295
         Begin VB.Frame Frame1 
            Caption         =   "Balloon Icon Type"
            Height          =   1335
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               HasDC           =   0   'False
               Height          =   975
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   1455
               TabIndex        =   24
               Top             =   240
               Width           =   1455
               Begin VB.OptionButton optIconType 
                  Caption         =   "Information"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   1
                  Left            =   120
                  TabIndex        =   28
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1245
               End
               Begin VB.OptionButton optIconType 
                  Caption         =   "Warning"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   2
                  Left            =   120
                  TabIndex        =   27
                  Top             =   480
                  Width           =   1365
               End
               Begin VB.OptionButton optIconType 
                  Caption         =   "Error"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   3
                  Left            =   120
                  TabIndex        =   26
                  Top             =   720
                  Width           =   945
               End
               Begin VB.OptionButton optIconType 
                  Caption         =   "None"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   25
                  Top             =   0
                  Width           =   795
               End
            End
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   6975
         TabIndex        =   18
         Top             =   2160
         Width           =   6975
         Begin VB.CheckBox chkCustomColors 
            Caption         =   "&Custom Colors"
            Height          =   195
            Left            =   5520
            TabIndex        =   31
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdTrayBalloonEX 
            Caption         =   "Show &Extended Tray Balloon"
            Height          =   375
            Left            =   2880
            TabIndex        =   29
            Top             =   0
            Width           =   2535
         End
         Begin VB.CommandButton cmdTrayBalloon 
            Caption         =   "&Show System Tray Balloon"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.TextBox txtTipText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmMain.frx":0000
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Text            =   "Hello!"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Timeout:"
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Balloon Prompt:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Balloon Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tray Icon Position in pixels:"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   4020
      Width           =   2895
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         ScaleHeight     =   375
         ScaleWidth      =   1575
         TabIndex        =   11
         Top             =   240
         Width           =   1575
         Begin VB.CommandButton cmdGetIconPos 
            Caption         =   "&Get"
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Label lblIconPos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblIconPos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblIconPos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblIconPos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Bottom:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Top:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Right:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Left:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show Tray Icon"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkVisible 
      Caption         =   "Tray Icon Visible"
      Height          =   195
      Left            =   1980
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Label lblDownloadComCtlDll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Download ""MS Common Controls DLL"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      MouseIcon       =   "frmMain.frx":0014
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Tag             =   "http://www.microsoft.com/downloads/details.aspx?displaylang=en&FamilyID=6F94D31A-D1E0-4658-A566-93AF0D8D4A1E"
      Top             =   3300
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "COMCTL32.DLL version:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Top             =   3300
      Width           =   1935
   End
   Begin VB.Label lblComctlVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2100
      TabIndex        =   35
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label lblVote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please, send your opinion and rate for me!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   60
      MouseIcon       =   "frmMain.frx":031E
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Tag             =   "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63280&lngWId=1"
      Top             =   6000
      Width           =   7275
   End
   Begin VB.Label lblSysTrayHWnd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2100
      TabIndex        =   21
      Top             =   3660
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "System Tray HWnd:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   3660
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuTrayLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayRate 
         Caption         =   "&Rate"
      End
      Begin VB.Menu mnuTrayLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const con_lToolTipCollor = &H80000018

Private Sub chkShow_Click()
  On Error GoTo lblErr
  
  If chkShow.Value = 1 Then
    With TrayIcon1
      'oTrayIcon.IconMovementTrackInterval = 500
      .TrayIconVisible = chkVisible.Value
      .IconHandle = Me.Icon
      .ToolTip = App.Title
      .Create Me.hwnd
    End With
  
  Else
    TrayIcon1.Remove
    ToolTipOnDemand1.Destroy
  End If

lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit
End Sub

Private Sub chkVisible_Click()
  On Error GoTo lblErr
  TrayIcon1.TrayIconVisible = chkVisible.Value

lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit
End Sub

Private Sub cmdGetIconPos_Click()
  On Error GoTo lblErr
  Dim rct As tiRECT
  rct = TrayIcon1.IconRect
  lblIconPos(0).Caption = rct.Left
  lblIconPos(1).Caption = rct.Right
  lblIconPos(2).Caption = rct.Top
  lblIconPos(3).Caption = rct.Bottom
  
lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit
End Sub

Private Sub cmdTrayBalloon_Click()
  On Error GoTo lblErr
  Dim i As Integer
  Dim enIconType As blIconType
  
  For i = 0 To optIconType.Count - 1
    If optIconType(i).Value Then
      enIconType = i
      Exit For
    End If
  Next i

  With TrayIcon1
    
    .BalloonTipShow enIconType, txtTipText.Text, txtTitle.Text, Val(txtTimeout.Text)
 
  End With

lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit
  
End Sub

Private Sub ShowBalloon(ByVal enIconType As blIconType, ByVal sPrompt As String, Optional ByVal sTitle As String, Optional ByVal lTimeout As Long, _
                        Optional ByVal lBackColor As Long = -1, Optional ByVal lForeColor As Long = -1)
  On Error GoTo lblErr
  Dim lX As Long, lY As Long
    
  Call TrayIcon1.GetIconMiddle(lX, lY)
  If lForeColor = -1 Then lForeColor = vbBlack
  If lBackColor = -1 Then lBackColor = &H80000018
  
  With ToolTipOnDemand1
    .ParentHwnd = TrayIcon1.SysTrayHWnd
    .x = lX
    .y = lY
    .BackColor = lBackColor
    .ForeColor = lForeColor
    .Prompt = sPrompt
    .Title = sTitle
    .TimeOut = lTimeout
    .IconType = enIconType
    .Show
  End With

lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit
End Sub

Private Sub cmdTrayBalloonEX_Click()
  On Error GoTo lblErr
  Dim i As Integer
  Dim enIconType As blIconType
  Dim lBackColor As Long, lForeColor As Long
  
  For i = 0 To optIconType.Count - 1
    If optIconType(i).Value Then
      enIconType = i
      Exit For
    End If
  Next i
    
    If chkCustomColors Then
      If enIconType = btNoIcon Then
        lBackColor = &HC0FFC0
        lForeColor = vbBlack
      ElseIf enIconType = btInfo Then
        lBackColor = &HC0FFC0
        lForeColor = vbBlack
      ElseIf enIconType = btWarning Then
        lBackColor = vbYellow
        lForeColor = vbBlack
      ElseIf enIconType = btError Then
        lBackColor = &HFF&
        lForeColor = vbWhite
      End If
    
    Else
      lBackColor = con_lToolTipCollor
      lForeColor = vbBlack
    End If
  
  Call ShowBalloon(enIconType, txtTipText.Text, txtTitle.Text, Val(txtTimeout.Text), lBackColor, lForeColor)
  
  If lblDownloadComCtlDll.Visible Then
    MsgBox "You need the updated MS Common Controls DLL", vbExclamation
    Call lblDownloadComCtlDll_Click
  End If
  
lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
   If Not TrayIcon1.Created Then
     Me.Enabled = False
     tmrExit.Enabled = True
   Else
    Call HideForm
   End If
  End If
  
End Sub

Private Sub Form_Load()
  On Error GoTo lblErr
  Dim sVersion() As String
  
  
  Me.Caption = App.Title
  
  With TrayIcon1
    .TrayIconVisible = chkVisible.Value
    .IconHandle = Me.Icon
    .ToolTip = App.Title
    .Create Me.hwnd
    lblComctlVersion.Caption = .CommonControlsVersion
    lblSysTrayHWnd.Caption = .SysTrayHWnd
  End With
  
  sVersion = Split(lblComctlVersion.Caption, ".")
  Call cmdGetIconPos_Click
  
  If Val(sVersion(0)) < 5 Then
    lblDownloadComCtlDll.Visible = True
  ElseIf Val(sVersion(0)) = 5 And Val(sVersion(1)) < 80 Then
    lblDownloadComCtlDll.Visible = True
  End If
  
lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    If TrayIcon1.Created Then
      Cancel = 1
      Call HideForm
    End If
  End If

End Sub

Private Sub HideForm()
  'Dim lX As Long, lY As Long
  Me.Hide
  Call ShowBalloon(btInfo, "I'm here", Me.Caption, 3000)
  
End Sub

Private Sub ShowForm()
  Dim lX As Long, lY As Long
  Me.Show
  SetForegroundWindow Me.hwnd
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo lblErr
  
  TrayIcon1.Remove
  ToolTipOnDemand1.Destroy
  
lblExit:
  Exit Sub
  
lblErr:
  MsgBox "Error #" & Err.Number & " in " & Err.Source & vbCrLf & Err.Description, vbCritical
  Resume lblExit
End Sub

Private Sub lblDownloadComCtlDll_Click()
  ShellExecute 0, "open", lblDownloadComCtlDll.Tag, vbNullString, vbNullString, 1&
End Sub

Private Sub lblVote_Click()
  ShellExecute 0, "open", lblVote.Tag, vbNullString, vbNullString, 1&
End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub mnuHide_Click()
  If Me.Visible Then
    Call HideForm
  Else
    ShowForm
  End If
End Sub

Private Sub mnuTrayExit_Click()
  Me.Enabled = False
  tmrExit.Enabled = True
End Sub

Private Sub mnuTrayRate_Click()
  Call lblVote_Click
End Sub

Private Sub ToolTipOnDemand1_BalloonDestroyed()
  TrayIcon1.TrackIconMovement = False
End Sub

Private Sub ToolTipOnDemand1_BalloonShowed()
  TrayIcon1.TrackIconMovement = True
End Sub

Private Sub TrayIcon1_BalloonClick(ByVal MouseEvent As stBalloonClickType)
  
  Select Case MouseEvent
    Case stbBalloonShow
      DispMsg "TrayIcon1_BalloonClick: stbBalloonShow"
    Case stbBalloonHide
      DispMsg "TrayIcon1_BalloonClick: stbBalloonHide"
    Case stbRightClick
      DispMsg "TrayIcon1_BalloonClick: stbRightClick"
    Case stbLeftClick
      DispMsg "TrayIcon1_BalloonClick: stbLeftClick"
    
  End Select

End Sub

Private Sub TrayIcon1_TaskBarRecreated()
  TrayIcon1.IconHandle = Me.Icon
  Call ShowBalloon(btInfo, "Windows Explorer has been restarted!", Me.Caption, 3000)
End Sub

Private Sub TrayIcon1_TrayIconMoved(ByVal lX As Long, ByVal lY As Long)
  Debug.Print "TrayIconMoved(" & lX & ", " & lY & ")"
  ToolTipOnDemand1.x = lX
  ToolTipOnDemand1.y = lY
  
End Sub

Private Sub TrayIcon1_TrayKeyEvent(ByVal KeyEvent As stKeyEvent)
  
  Select Case KeyEvent
    Case stSelect
      DispMsg "TrayIcon1_TrayKeyEvent: stSelect"
    Case stKeySelect
      DispMsg "TrayIcon1_TrayKeyEvent: stKeySelect"
    
  End Select

End Sub

Private Sub TrayIcon1_TrayMouseEvent(ByVal MouseEvent As stMouseEvent)
  
  Select Case MouseEvent
    Case stMouseMove
      DispMsg "TrayIcon1_TrayMouseEvent: stMouseMove"
    
    Case stLeftButtonDown
      SetForegroundWindow Me.hwnd
      DispMsg "TrayIcon1_TrayMouseEvent: stLeftButtonDown"
      
    Case stLeftButtonUp
      DispMsg "TrayIcon1_TrayMouseEvent: stLeftButtonUp"
      
    Case stLeftButtonDoubleClick
      DispMsg "TrayIcon1_TrayMouseEvent: stLeftButtonDoubleClick"
      Call mnuHide_Click
      
    Case stRightButtonDown
      DispMsg "TrayIcon1_TrayMouseEvent: stRightButtonDown"
      If Me.Visible Then
        mnuHide.Caption = "&Hide"
        SetForegroundWindow Me.hwnd
      Else
        mnuHide.Caption = "&Restore"
      End If
      PopupMenu mnuTray
      
    Case stRightButtonUp
      DispMsg "TrayIcon1_TrayMouseEvent: stRightButtonUp"
      
    Case stRightButtonDoubleClick
      DispMsg "TrayIcon1_TrayMouseEvent: stRightButtonDoubleClick"
      
    Case stMiddleButtonDown
      DispMsg "TrayIcon1_TrayMouseEvent: stMiddleButtonDown"
      
    Case stMiddleButtonUp
      DispMsg "TrayIcon1_TrayMouseEvent: stMiddleButtonUp"
      
    Case stMiddleButtonDoubleClick
      DispMsg "TrayIcon1_TrayMouseEvent: stMiddleButtonDoubleClick"
    
  End Select

End Sub

Private Sub ToolTipOnDemand1_MouseEvents(MouseEvent As Long)
  
  Select Case MouseEvent
    Case stMouseMove
      DispMsg "ToolTipOnDemand1_MouseEvents: stMouseMove"
    
    Case stLeftButtonDown
      DispMsg "ToolTipOnDemand1_MouseEvents: stLeftButtonDown"
      ToolTipOnDemand1.Destroy
      
    Case stLeftButtonUp
      DispMsg "ToolTipOnDemand1_MouseEvents: stLeftButtonUp"
      
    Case stLeftButtonDoubleClick
      DispMsg "ToolTipOnDemand1_MouseEvents: stLeftButtonDoubleClick"
      
    Case stRightButtonDown
      ToolTipOnDemand1.Destroy
      DispMsg "ToolTipOnDemand1_MouseEvents: stRightButtonDown"
      
    Case stRightButtonUp
      DispMsg "ToolTipOnDemand1_MouseEvents: stRightButtonUp"
      
    Case stRightButtonDoubleClick
      DispMsg "ToolTipOnDemand1_MouseEvents: stRightButtonDoubleClick"
      
    Case stMiddleButtonDown
      DispMsg "ToolTipOnDemand1_MouseEvents: stMiddleButtonDown"
      
    Case stMiddleButtonUp
      DispMsg "ToolTipOnDemand1_MouseEvents: stMiddleButtonUp"
      
    Case stMiddleButtonDoubleClick
      DispMsg "ToolTipOnDemand1_MouseEvents: stMiddleButtonDoubleClick"
    
  End Select

End Sub

Private Sub DispMsg(sMsg As String)
  Static iErrCount As Integer
  On Error GoTo lblErr
  txtEvents.SelStart = Len(txtEvents.Text)
  txtEvents.SelText = sMsg & vbCrLf
  
  
lblExit:
  Exit Sub
  
lblErr:
  Debug.Print "Error #" & Err.Number & " in " & Err.Source & ", " & Err.Description, vbCritical
  If Err.Number = 380 Then
    If iErrCount = 1 Then
      iErrCount = 0
      Resume lblExit
    Else
      iErrCount = 1
      txtEvents.Text = ""
      Resume 0
    End If
  End If
  
End Sub

Private Sub tmrExit_Timer()
  tmrExit.Enabled = False
  Unload Me
End Sub
