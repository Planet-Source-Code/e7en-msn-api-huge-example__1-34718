Attribute VB_Name = "Mouse"
'**************************************
' Name: A mouse module, FINALLY!!! Move,
'     click, +more
' Description:This module has the follow
'     ing functions (pretty self explanitory):
'     GetX, GetY, LeftClick, LeftDown, LeftUp,
'     RightClick, RightUp, RightDown, MiddleCl
'     ick, MiddleDown, MiddleUp, MoveMouse, Se
'     tMousePos
' By: Arthur Chaparyan
'
' Assumes:You should know how to create
'     and use a module. If you have any questi
'     ons, please submit a comment, thanX
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.2795/lngWId.1/qx/vb/scripts/ShowCode.
'     htm'for details.'**************************************

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Const MOUSEEVENTF_LEFTDOWN = 2
    Const MOUSEEVENTF_LEFTUP = 4
    Const MOUSEEVENTF_MIDDLEDOWN = 32
    Const MOUSEEVENTF_MIDDLEUP = 64
    Const MOUSEEVENTF_MOVE = 1
    Const MOUSEEVENTF_RIGHTDOWN = 8
    Const MOUSEEVENTF_RIGHTUP = 16


Public Type POINTAPI
    x As Long
    y As Long
    End Type




Public Function GetX() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetX = n.x
End Function


Public Function GetY() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetY = n.y
End Function


Public Sub LeftClick()
    LeftDown
    LeftUp
End Sub


Public Sub LeftDown()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub


Public Sub LeftUp()
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub


Public Sub MiddleClick()
    MiddleDown
    MiddleUp
End Sub


Public Sub MiddleDown()
    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
End Sub


Public Sub MiddleUp()
    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
End Sub


Public Sub MoveMouse(xMove As Long, yMove As Long)
    mouse_event MOUSEEVENTF_MOVE, xMove, yMove, 0, 0
End Sub


Public Sub RightClick()
    RightDown
    RightUp
End Sub


Public Sub RightDown()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub


Public Sub RightUp()
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub


Public Sub SetMousePos(xPos As Long, yPos As Long)
    SetCursorPos xPos, yPos
End Sub

