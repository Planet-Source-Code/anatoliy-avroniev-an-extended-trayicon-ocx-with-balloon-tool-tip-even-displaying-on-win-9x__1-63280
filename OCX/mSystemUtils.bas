Attribute VB_Name = "mSystemUtils"
Option Explicit
' ======================================================================================
' Name:     mSystemUtils module
' Author:   Anatoliy Avroniev (aavroniev.axenet.ru)
' Date:     07 December 2004
' ======================================================================================

Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Const FORMAT_MESSAGE_FROM_STRING = &H400
Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Private Declare Function FormatMessage Lib "kernel32" Alias _
   "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
   ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
   ByVal lpBuffer As String, ByVal nSize As Long, _
   Arguments As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Enum enPlatformID
  VER_PLATFORM_WIN32s = 0          'WINDOWS 3.1
  VER_PLATFORM_WIN32_WINDOWS = 1   'WINDOWS 9x
  VER_PLATFORM_WIN32_NT = 2        'WINDOWS NT, 2000, XP
End Enum

Public Enum enProductType
  VER_NT_WORKSTATION = &H1
  VER_NT_DOMAIN_CONTROLLER = &H2
  VER_NT_SERVER = &H3
End Enum

Public Enum enSuiteMask
  VER_SUITE_SMALLBUSINESS = &H1
  VER_SUITE_ENTERPRISE = &H2
  VER_SUITE_BACKOFFICE = &H4
  VER_SUITE_COMMUNICATIONS = &H8
  VER_SUITE_TERMINAL = &H10
  VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20
  VER_SUITE_EMBEDDEDNT = &H40
  VER_SUITE_DATACENTER = &H80
  VER_SUITE_SINGLEUSERTS = &H100
  VER_SUITE_PERSONAL = &H200
  VER_SUITE_BLADE = &H400
End Enum

Public Type udtOSVERSIONINFO
    lMajorVersion As Long
    lMinorVersion As Long
    lBuildNumber As Long
    lPlatformId As enPlatformID
    sCSDVersion As String
    iServicePackMajor As Integer
    iServicePackMinor As Integer
    iSuiteMask As enSuiteMask
    bytProductType As enProductType
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
  wServicePackMajor As Integer
  wServicePackMinor As Integer
  wSuiteMask As Long
  wProductType As Byte
  wReserved As Byte
End Type

Private Declare Function GetVersion Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private m_pUdtOSVersion As OSVERSIONINFO
Private m_pUdtOSVersionEX As OSVERSIONINFOEX

Public tOSVERSIONINFO As udtOSVERSIONINFO

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Public Function APIErrorDescription(ErrorCode As Long) As String

Dim sAns As String
Dim lRet As Long

'PURPOSE: Returns Human Readable Description of
'Error Code that occurs in API function

'PARAMETERS: ErrorCode: System Error Code

'Returns: Description of Error

sAns = Space(255)
lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, _
  ErrorCode, 0, sAns, 255, 0)

APIErrorDescription = fNullCut(sAns)

End Function

Public Sub GetOSVersion()
  On Error GoTo ErrorHandler
  Dim pUdtOSVersion As OSVERSIONINFO
  Dim pUdtOSVersionEX As OSVERSIONINFOEX
  Dim pOSVERSIONINFO As udtOSVERSIONINFO
  Dim plMajorVersion  As Long
  Dim plMinorVersion As Long
  Dim plPlatformID As Long
  Dim sCSDVer As String
  Dim lRetVal As Long
    
    m_pUdtOSVersion = pUdtOSVersion
    m_pUdtOSVersionEX = pUdtOSVersionEX
    
    pUdtOSVersion.dwOSVersionInfoSize = Len(pUdtOSVersion)
    lRetVal = GetVersion(pUdtOSVersion)
    If lRetVal = 0 Then GoTo lblEnd
    m_pUdtOSVersion = pUdtOSVersion
    
    plMajorVersion = pUdtOSVersion.dwMajorVersion
    plMinorVersion = pUdtOSVersion.dwMinorVersion
    plPlatformID = pUdtOSVersion.dwPlatformId
    sCSDVer = UCase(pUdtOSVersion.szCSDVersion)
    
    If plMajorVersion = 5 Then
      pUdtOSVersionEX.dwOSVersionInfoSize = Len(pUdtOSVersionEX)
      lRetVal = GetVersionEx(pUdtOSVersionEX)
      If lRetVal <> 0 Then m_pUdtOSVersionEX = pUdtOSVersionEX
    End If

With pOSVERSIONINFO
  .lBuildNumber = pUdtOSVersion.dwBuildNumber
  .lMajorVersion = pUdtOSVersion.dwMajorVersion
  .lMinorVersion = pUdtOSVersion.dwMinorVersion
  .lPlatformId = pUdtOSVersion.dwPlatformId
  .sCSDVersion = fNullCut(pUdtOSVersion.szCSDVersion)
  .iServicePackMajor = pUdtOSVersionEX.wServicePackMajor
  .iServicePackMinor = pUdtOSVersionEX.wServicePackMinor
  .iSuiteMask = pUdtOSVersionEX.wSuiteMask
  .bytProductType = pUdtOSVersionEX.wProductType
End With


lblEnd:
tOSVERSIONINFO = pOSVERSIONINFO

lblExit:
  Exit Sub

ErrorHandler:
  tOSVERSIONINFO = pOSVERSIONINFO
  Resume lblExit
  
End Sub

Private Function fLoWord(ByVal dwValue As Long) As Long
  fLoWord = (dwValue And &HFFFF&)
End Function

Private Function fNullCut(ByVal myString As String) As String
  Dim i As Long
  i = InStr(myString, vbNullChar)
  If i > 0& Then
    fNullCut = Trim$(Left$(myString, i - 1&))
  Else
    fNullCut = Trim$(myString)
  End If
End Function

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


