VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************************************************
'GSubclass
'  Credits:
'    Steve McMahon: http://www.vbaccelerator.com/home/VB/Code/Libraries/Subclassing/SSubTimer/article.asp
'**********************************************************************************************************

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Function IsWindowLocal(ByVal hwnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hwnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function
'Return whether we're running in the IDE. Public for general utility purposes
Public Function InIDE() As Boolean
    Debug.Assert SetTrue(InIDE)
End Function
'Worker function for InIDE - will only be called whilst running in the IDE
Private Function SetTrue(bValue As Boolean) As Boolean
    SetTrue = True
    bValue = True
End Function

Sub AttachMessage(iwp As ISubclass, ByVal hwnd As Long, _
                  ByVal iMsg As Long)
Attribute AttachMessage.VB_Description = "Start subclassing for a particular message."
    MSubclass.AttachMessage iwp, hwnd, iMsg
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hwnd As Long, _
                  ByVal iMsg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
    MSubclass.DetachMessage iwp, hwnd, iMsg
End Sub

Public Property Get CurrentMessage() As Long
Attribute CurrentMessage.VB_Description = "Returns the message currently being worked on.  Only valid in the MsgResponse and WindowProc items of your implemented interface."
   CurrentMessage = MSubclass.CurrentMessage
End Property
Public Function CallOldWindowProc( _
      ByVal hwnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
Attribute CallOldWindowProc.VB_Description = "Calls the window procedure for this handle which was installed before you added the subclass."
   CallOldWindowProc = MSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
End Function

