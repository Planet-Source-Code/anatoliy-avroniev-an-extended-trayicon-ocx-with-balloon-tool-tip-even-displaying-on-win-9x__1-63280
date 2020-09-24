Attribute VB_Name = "mTray"
Option Explicit
' ======================================================================================
' Name:     mTray module
' Author:   Anatoliy Avroniev (aavroniev.axenet.ru)
' Date:     15 May 2006
' ======================================================================================

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const IMAGE_BITMAP = 0&
Private Const IMAGE_ICON = 1&
Private Const IMAGE_CURSOR = 2&

Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBHeader = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000

Private Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hwnd        As Long
    uID         As Long
    cRect       As RECT
    hInst       As Long
    lpszText    As Long 'LPCSTR
    lParam      As Long
End Type

Private Type TOOLTEXT
    sTipText As String * 80
End Type

Public Type udtTrayList
  trTipText As String
  trToolInfo As TOOLINFO
End Type

Private Const WM_USER = &H400
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_GETTEXTW = (WM_USER + 56)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Const GWL_STYLE = (-16)
Private Const GWL_HINSTANCE = (-6)
Private Const TTS_NOPREFIX = 2


Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private NWPid As Long
Private tTrayList() As udtTrayList

Private Sub GetTrayList()
   Dim NWThreadID As Long
   
   ReDim tTrayList(0)
   NWThreadID = 0
   NWPid = 0
   NWThreadID = GetWindowThreadProcessId(GetTrayNotifyWnd, NWPid)
   EnumWindows AddressOf EnumWinProc, 0
   
End Sub

Private Function GetTrayNotifyWnd() As Long
   GetTrayNotifyWnd = FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString)
End Function

Private Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
   Dim pid As Long, tid As Long, lStyle As Long
   Dim hProcess As Long, nCount As Long, lWritten As Long, i As Long
   Dim lpSysShared As Long, hFileMapping As Long, dwSize As Long
   Dim lpSysShared2 As Long, hFileMapping2 As Long
   Dim h As Long
   Dim ti As TOOLINFO
   Dim tt As TOOLTEXT
   Dim rc As RECT
   Static sTip As String
   
   tid = GetWindowThreadProcessId(hwnd, pid)
   lStyle = GetWindowLong(hwnd, GWL_STYLE)
   
   If pid = NWPid And GetWndClass(hwnd) = "tooltips_class32" And (lStyle And TTS_NOPREFIX) <> TTS_NOPREFIX Then
      nCount = SendMessage(hwnd, TTM_GETTOOLCOUNT, 0&, ByVal 0&)
      If nCount <= 0 Then
        EnumWinProc = 0
        Exit Function
      End If
      
      ReDim tTrayList(nCount - 1)
      'ReDim sTips(nCount - 1)
      'ReDim m_rcIconRects(nCount - 1)
      ti.cbSize = Len(ti)
      dwSize = ti.cbSize
      
      If IsWindowsNT Then
         lpSysShared = GetMemSharedNT(pid, dwSize, hProcess)
         If lpSysShared = 0 Then GoTo lblEnd
         lpSysShared2 = GetMemSharedNT(pid, LenB(tt), hProcess)
         If lpSysShared2 = 0 Then GoTo lblEnd
         WriteProcessMemory hProcess, ByVal lpSysShared, ti, dwSize, lWritten
         
         For i = 0 To nCount - 1
             tt.sTipText = String(80, Chr(0))
             WriteProcessMemory hProcess, ByVal lpSysShared2, tt, LenB(tt), lWritten
             Call SendMessage(hwnd, TTM_ENUMTOOLSW, i, ByVal lpSysShared)
             ReadProcessMemory hProcess, ByVal lpSysShared, ti, dwSize, lWritten
             ti.lpszText = lpSysShared2
             WriteProcessMemory hProcess, ByVal lpSysShared, ti, dwSize, lWritten
             Call SendMessage(hwnd, TTM_GETTEXTW, 0&, ByVal lpSysShared)
             ReadProcessMemory hProcess, ByVal lpSysShared2, tt, LenB(tt), lWritten
             
             tTrayList(i).trToolInfo = ti
             
             sTip = StrConv(tt.sTipText, vbFromUnicode)
             sTip = StripTerminator(sTip)
             'Debug.Print "sTip=" & sTip & ", hwnd=" & hwnd
             tTrayList(i).trTipText = sTip
             If i = nCount - 1 Then Exit For
         Next i
         
         FreeMemSharedNT hProcess, lpSysShared, dwSize
         FreeMemSharedNT hProcess, lpSysShared2, LenB(tt)
      
      Else
         lpSysShared = GetMemShared95(dwSize, hFileMapping)
         If lpSysShared = 0 Then GoTo lblEnd
         lpSysShared2 = GetMemShared95(Len(tt), hFileMapping2)
         If lpSysShared2 = 0 Then GoTo lblEnd
         
         CopyMemory ByVal lpSysShared, ti, dwSize
         For i = 0 To nCount - 1
             tt.sTipText = String(80, Chr(0))
             CopyMemory ByVal lpSysShared2, tt, Len(tt)
             Call SendMessage(hwnd, TTM_ENUMTOOLSA, i, ByVal lpSysShared)
             CopyMemory ti, ByVal lpSysShared, dwSize
             ti.lpszText = lpSysShared2
             CopyMemory ByVal lpSysShared, ti, dwSize
             Call SendMessage(hwnd, TTM_GETTEXTA, 0&, ByVal lpSysShared)
             CopyMemory tt, ByVal lpSysShared2, Len(tt)
             
             tTrayList(i).trToolInfo = ti
             sTip = StripTerminator(tt.sTipText)
             'Debug.Print "sTip=" & sTip & ", hwnd=" & hwnd
             tTrayList(i).trTipText = sTip
             If i = nCount - 1 Then Exit For
         Next i
         FreeMemShared95 hFileMapping, lpSysShared
         FreeMemShared95 hFileMapping2, lpSysShared2
      End If
      
lblEnd:
      EnumWinProc = 0
      Exit Function
   
   End If
   EnumWinProc = 1
End Function

Private Function GetWndClass(hwnd As Long) As String
   Dim k As Long, sName As String
   sName = Space$(128)
   k = GetClassName(hwnd, sName, 128)
   If k > 0 Then sName = Left$(sName, k) Else sName = ""
   GetWndClass = sName
End Function

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos = 1 Then
        StripTerminator = ""
    ElseIf intZeroPos > 1 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function GetSysTrayHWnd() As Long
    Const GW_CHILD = 5
    Const GW_HWNDNEXT = 2
    '
    Dim lngTaskbarHwnd As Long
    Dim lHWnd As Long
    Dim lTrayHwnd As Long
    Dim strClassName As String * 250
    Dim sClassName As String
    Dim hInstance As Long
    Dim rc As RECT
    Dim n As Long, hIcon As Long
    
    '
    'Get taskbar handle
    lngTaskbarHwnd = FindWindow("Shell_TrayWnd", vbNullString)
    If lngTaskbarHwnd = 0 Then GoTo lblExit
    '
    'Get system tray handle
    'Call PirntText("-------TrayNotifyWnd-------")
    lHWnd = GetWindow(lngTaskbarHwnd, GW_CHILD)
    If lHWnd = 0 Then GoTo lblExit
    
    Do Until lHWnd = 0
        strClassName = String(250, Chr(0))
        GetClassName lHWnd, strClassName, 250
        sClassName = StripTerminator(strClassName)
        'Call PirntText("sClassName=" & sClassName & ", Hwnd=" & lHwnd)
        If sClassName = "TrayNotifyWnd" Then
          lTrayHwnd = lHWnd
          Exit Do
        End If
        lHWnd = GetWindow(lHWnd, GW_HWNDNEXT)
    Loop
    '
    'Call PirntText("-------SysPager-------")
    lHWnd = GetWindow(lTrayHwnd, GW_CHILD)
    If lHWnd = 0 Then GoTo lblExit
    
    Do Until lHWnd = 0
        strClassName = String(250, Chr(0))
        GetClassName lHWnd, strClassName, 250
        sClassName = StripTerminator(strClassName)
        'Call PirntText("sClassName=" & sClassName & ", Hwnd=" & lHwnd)
        If sClassName = "SysPager" Then
          lTrayHwnd = lHWnd
          Exit Do
        End If
        lHWnd = GetWindow(lHWnd, GW_HWNDNEXT)
    Loop
    
    'Call PirntText("-------ToolbarWindow32-------")
    lHWnd = GetWindow(lTrayHwnd, GW_CHILD)
    If lHWnd = 0 Then GoTo lblExit
    
    Do Until lHWnd = 0
        strClassName = String(250, Chr(0))
        GetClassName lHWnd, strClassName, 250
        sClassName = StripTerminator(strClassName)
        'Call PirntText("sClassName=" & sClassName & ", Hwnd=" & lHwnd)
        If sClassName = "ToolbarWindow32" Then
          lTrayHwnd = lHWnd
          Exit Do
        End If
        lHWnd = GetWindow(lHWnd, GW_HWNDNEXT)
    Loop
    

lblExit:
    GetSysTrayHWnd = lTrayHwnd

End Function

Private Function GetToolTipHandle() As Long
  Dim lRet  As Long
  Dim lTaskBar As Long
  Dim pidTaskBar As Long
  Dim wnd As Long
  Dim pidWnd As Long
  
  'Get the TaskBar handle
  lTaskBar = FindWindowEx(0, 0, "Shell_TrayWnd", vbNullString)
  If lTaskBar = 0 Then Exit Function
  
  'Get the TaskBar Process ID
  lRet = GetWindowThreadProcessId(lTaskBar, pidTaskBar)
  If pidTaskBar = 0 Then Exit Function
  
  'Enumerate all tooltip windows
  wnd = FindWindowEx(0&, 0&, "tooltips_class32", vbNullString)
  Do While wnd <> 0
    'Get the tooltip process ID
    lRet = GetWindowThreadProcessId(wnd, pidWnd)
    
    'Compare the process ID of the taskbar and the tooltip.
    'If they are the same we have one of the taskbar tooltips.
    If pidTaskBar = pidWnd Then
      'Get the tooltip style. The tooltip for tray icons does not have the
      'TTS_NOPREFIX style.
      If (GetWindowLong(wnd, GWL_STYLE) And TTS_NOPREFIX) = 0 Then
        Exit Do
      End If
    End If
    wnd = FindWindowEx(0, wnd, "tooltips_class32", vbNullString)
    
  Loop

  GetToolTipHandle = wnd
End Function

Public Function GetTrayIconRect() As RECT
    Dim sTrayListArray() As String
    Dim rctSysTray As RECT
    Dim rctTrayIcon As RECT
    Dim i As Integer, bIconFound As Boolean
    Dim SysTrayHWnd As Long
    
  
    SysTrayHWnd = GetSysTrayHWnd()
    If SysTrayHWnd = 0 Then GoTo lblExit
    Call GetWindowRect(SysTrayHWnd, rctSysTray)
    
    Call GetTrayList
    If UBound(tTrayList) = 0 And tTrayList(0).trToolInfo.hwnd = 0 Then GoTo lblExit
    
    For i = 0 To UBound(tTrayList)
      'Debug.Print i & "-'" & tTrayList(i).trTipText & "'"
      If tTrayList(i).trTipText = "wAnTeD tRaY iCoN " & App.hInstance Then
        bIconFound = True
        Exit For
      End If
    Next i
    
    If bIconFound Then
      With tTrayList(i).trToolInfo.cRect
        rctTrayIcon.Left = rctSysTray.Left + .Left
        rctTrayIcon.Right = rctSysTray.Left + .Right
        rctTrayIcon.Top = rctSysTray.Top + .Top
        rctTrayIcon.Bottom = rctSysTray.Top + .Bottom
      End With
    
      GetTrayIconRect = rctTrayIcon
    End If
    
lblExit:

End Function
