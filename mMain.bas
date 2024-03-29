Attribute VB_Name = "mMain"
Option Explicit
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200


Public Sub Main()
  Dim b As Boolean
  On Error Resume Next
  
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   b = InitCommonControlsEx(iccex)
   
   On Error GoTo 0
   frmMain.Show
   
lblExit:
  Exit Sub
  
lblEnd:
  End
End Sub

