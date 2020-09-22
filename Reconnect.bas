Attribute VB_Name = "Module1"
Global Const VK_SPACE = &H20
Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101

Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetTickCount& Lib "kernel32" ()
' Declare Function FindChildByTitle% Lib "vbwfind.dll" (ByVal parent%, ByVal title$)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal hWnd As Integer) As Integer


Declare Function SetActiveWindow Lib "User" (ByVal hWnd As Integer) As Integer


Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


    'Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA"
    '     (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String,
    '     ByVal lpsz2 As String) As Long
    'find child by class...


Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long


Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Declare Function GetNextWindow Lib "User32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
    'find child by title...


Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

