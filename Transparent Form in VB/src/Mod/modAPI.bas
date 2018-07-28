Attribute VB_Name = "ModAPi"
'//This is Module
Option Explicit

Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" _
   (ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
