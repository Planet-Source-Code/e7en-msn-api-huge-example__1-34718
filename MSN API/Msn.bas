Attribute VB_Name = "Msn"
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type


Public Type TVHITTESTINFO
    pt As POINTAPI
    flags As Long
    hItem As Long
End Type

Public prevItem As Node
Global TVhwnd As Long
Public Const GWL_STYLE = (-16)
Public Const TV_FIRST = &H1100
Public Const TVM_HITTEST = (TV_FIRST + 17)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Function TreeView_HitTest(hWnd As Long, lpht As TVHITTESTINFO) As Long
  
    TreeView_HitTest = SendMessage(hWnd, TVM_HITTEST, 0, lpht)

End Function
