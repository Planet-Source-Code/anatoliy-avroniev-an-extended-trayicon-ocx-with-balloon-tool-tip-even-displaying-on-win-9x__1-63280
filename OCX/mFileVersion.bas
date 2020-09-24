Attribute VB_Name = "mFileVersion"
Option Explicit

'  ----- Symbols -----
Const VS_VERSION_INFO = 1
Const VS_USER_DEFINED = 100

'  ----- VS_VERSION.dwFileFlags -----
Const VS_FFI_SIGNATURE = &HFEEF04BD
Const VS_FFI_STRUCVERSION = &H10000
Const VS_FFI_FILEFLAGSMASK = &H3F&

'  ----- VS_VERSION.dwFileFlags -----
Const VS_FF_DEBUG = &H1&
Const VS_FF_PRERELEASE = &H2&
Const VS_FF_PATCHED = &H4&
Const VS_FF_PRIVATEBUILD = &H8&
Const VS_FF_INFOINFERRED = &H10&
Const VS_FF_SPECIALBUILD = &H20&

'  ----- VS_VERSION.dwFileOS -----
Const VOS_UNKNOWN = &H0&
Const VOS_DOS = &H10000
Const VOS_OS216 = &H20000
Const VOS_OS232 = &H30000
Const VOS_NT = &H40000

Const VOS__BASE = &H0&
Const VOS__WINDOWS16 = &H1&
Const VOS__PM16 = &H2&
Const VOS__PM32 = &H3&
Const VOS__WINDOWS32 = &H4&

Const VOS_DOS_WINDOWS16 = &H10001
Const VOS_DOS_WINDOWS32 = &H10004
Const VOS_OS216_PM16 = &H20002
Const VOS_OS232_PM32 = &H30003
Const VOS_NT_WINDOWS32 = &H40004

'  ----- VS_VERSION.dwFileType -----
Const VFT_UNKNOWN = &H0&
Const VFT_APP = &H1&
Const VFT_DLL = &H2&
Const VFT_DRV = &H3&
Const VFT_FONT = &H4&
Const VFT_VXD = &H5&
Const VFT_STATIC_LIB = &H7&

'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Const VFT2_UNKNOWN = &H0&
Const VFT2_DRV_PRINTER = &H1&
Const VFT2_DRV_KEYBOARD = &H2&
Const VFT2_DRV_LANGUAGE = &H3&
Const VFT2_DRV_DISPLAY = &H4&
Const VFT2_DRV_MOUSE = &H5&
Const VFT2_DRV_NETWORK = &H6&
Const VFT2_DRV_SYSTEM = &H7&
Const VFT2_DRV_INSTALLABLE = &H8&
Const VFT2_DRV_SOUND = &H9&
Const VFT2_DRV_COMM = &HA&
Const VFT2_DRV_INPUTMETHOD = &HB&

'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_FONT -----
Const VFT2_FONT_RASTER = &H1&
Const VFT2_FONT_VECTOR = &H2&
Const VFT2_FONT_TRUETYPE = &H3&

'  ----- VerFindFile() flags -----
Const VFFF_ISSHAREDFILE = &H1

Const VFF_CURNEDEST = &H1
Const VFF_FILEINUSE = &H2
Const VFF_BUFFTOOSMALL = &H4

'  ----- VerInstallFile() flags -----
Const VIFF_FORCEINSTALL = &H1
Const VIFF_DONTDELETEOLD = &H2

Const VIF_TEMPFILE = &H1&
Const VIF_MISMATCH = &H2&
Const VIF_SRCOLD = &H4&

Const VIF_DIFFLANG = &H8&
Const VIF_DIFFCODEPG = &H10&
Const VIF_DIFFTYPE = &H20&

Const VIF_WRITEPROT = &H40&
Const VIF_FILEINUSE = &H80&
Const VIF_OUTOFSPACE = &H100&
Const VIF_ACCESSVIOLATION = &H200&
Const VIF_SHARINGVIOLATION = &H400&
Const VIF_CANNOTCREATE = &H800&
Const VIF_CANNOTDELETE = &H1000&
Const VIF_CANNOTRENAME = &H2000&
Const VIF_CANNOTDELETECUR = &H4000&
Const VIF_OUTOFMEMORY = &H8000&

Const VIF_CANNOTREADSRC = &H10000
Const VIF_CANNOTREADDST = &H20000
Const VIF_BUFFTOOSMALL = &H40000

'  ----- Types and structures -----

Private Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type

'  ----- Function prototypes -----

Private Declare Function VerFindFile Lib "version.dll" Alias "VerFindFileA" (ByVal uFlags As Long, ByVal szFileName As String, ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, lpuCurDirLen As Long, ByVal szDestDir As String, lpuDestDirLen As Long) As Long
Private Declare Function VerInstallFile Lib "version.dll" Alias " VerInstallFileA" (ByVal uFlags As Long, ByVal szSrcFileName As String, ByVal szDestFileName As String, ByVal szSrcDir As String, ByVal szDestDir As String, ByVal szCurDir As String, ByVal szTmpFile As String, lpuTmpFileLen As Long) As Long

'  Returns size of version info in Bytes
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

'  Read version info into buffer
' /* Length of buffer for info *
' /* Information from GetFileVersionSize *
' /* Filename of version stamped file *
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function FileVersion(ByVal sFile As String) As String

   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   Dim sBuffer As String
   Dim tFixedFileInfo As VS_FIXEDFILEINFO
   Dim i As Integer, sVer As String
   Dim lRet As Long
   
   
   nBufferSize = GetFileVersionInfoSize(sFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sFile, 0&, nBufferSize, bBuffer(0))
      
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory tFixedFileInfo, ByVal lpBuffer, Len(tFixedFileInfo)
         
         If tFixedFileInfo.dwFileVersionMS > 0 Then
         
           sVer = HiWord(tFixedFileInfo.dwFileVersionMS) & "." & LoWord(tFixedFileInfo.dwFileVersionMS) & "." _
               & HiWord(tFixedFileInfo.dwFileVersionLS) & "." & LoWord(tFixedFileInfo.dwFileVersionLS)
         
         End If
         
         FileVersion = sVer
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

Private Function HiWord(ByVal wParam As Long) As Integer
   HiWord = (wParam And &HFFFF0000) \ (&H10000)

End Function

Private Function LoWord(ByVal wParam As Long) As Integer
   LoWord = wParam And &HFFFF&

End Function

Private Function HiByte(ByVal wParam As Integer) As Byte
   
   HiByte = (wParam And &HFF00&) \ (&H100)

End Function

Private Function LoByte(ByVal wParam As Integer) As Byte

   LoByte = wParam And &HFF&

End Function

