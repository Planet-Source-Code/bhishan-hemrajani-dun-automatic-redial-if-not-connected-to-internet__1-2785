Attribute VB_Name = "basMain"
Option Explicit
'-------------------------------------------------------------------------------------
'WIN32 API Constants
'-------------------------------------------------------------------------------------
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_WANTMAPPINGHANDLE = &H20
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_FILESONLY = &H80
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_ALLOWUNDO = &H40
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GWL_WNDPROC = (-4)
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_USER = &H400

'-------------------------------------------------------------------------------------
'Application Defined Constants
'-------------------------------------------------------------------------------------
Public Const OPT_COPY = 0
Public Const OPT_MOVE = 1
Public Const OPT_DELETE = 2
Public Const OPT_RENAME = 3
Public Const WM_CALLBACK_MSG = WM_USER Or &HF

'-------------------------------------------------------------------------------------
'WIN32 Type Declares
'-------------------------------------------------------------------------------------
Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'-------------------------------------------------------------------------------------
'Application defined enumerations
'-------------------------------------------------------------------------------------
Public Enum ZoomTypes
    ZOOM_FROM_TRAY
    ZOOM_TO_TRAY
End Enum

'-------------------------------------------------------------------------------------
'WIN32 DLL Declares
'-------------------------------------------------------------------------------------
Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Declare Function SHFileOperationA Lib "shell32.dll" (lpFileOp As Any) As Long

Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                                         ByVal lpWindowName As String) As Long

Declare Function GetDesktopWindow Lib "User32" () As Long

Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, _
                                                                                                ByVal lpClassName As String, _
                                                                                                ByVal nMaxCount As Long) As Long

Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long

Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Declare Function DrawAnimatedRects Lib "User32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, _
                                                                       lprcTo As RECT) As Long

Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                                                   ByVal hWnd As Long, _
                                                                                                   ByVal Msg As Long, _
                                                                                                   ByVal wParam As Long, _
                                                                                                   ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                                                    ByVal nIndex As Long, _
                                                                                                    ByVal dwNewLong As Long) As Long
                                                                                                    
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                                             ByVal wMsg As Long, _
                                                                                             ByVal wParam As Long, _
                                                                                             lParam As Any) As Long
                                                                                             
Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, _
                                                                                                           lpData As NOTIFYICONDATA) _
                                                                                                           As Long
                                                                                                           
Public lngPrevWndProc As Long 'Original WNDPROC address.
Public lngWndID As Long 'Our unique icon identifier.
Public lngHwnd As Long 'The hwnd of frmMain.
Public nidTray As NOTIFYICONDATA 'Global Icon Info Structure.
Public strToolTip As String 'The string we use in the tip text fro the tray.
Public intOptionChoice As Integer 'Option button tracker.

Public Function WndProcMain(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long
'-------------------------------------------------------------------------------------
'Window Proc for frmMain.  Used for sub-classing.
'-------------------------------------------------------------------------------------
If lngWndID = wParam Then
    Select Case lParam
        Case WM_LBUTTONDBLCLK
            'If the user double clicked the tray icon then release the hook on the window.
            'In other words un-subclass it by returning it's original WNDPROC address
            'that we captured in the original call to SetWindowLong.
            SetWindowLong lngHwnd, GWL_WNDPROC, lngPrevWndProc
            
            'Delete our icon from the tray.
            Shell_NotifyIconA NIM_DELETE, nidTray
            ZoomForm ZOOM_FROM_TRAY, frmMain.hWnd
            DoEvents
            frmMain.Show
            DoEvents
        Case WM_RBUTTONDOWN
            'Show our popup menu in the tray.
            frmMain.PopupMenu frmMain.mnuPopUp

    End Select
End If

'Call the original WNDPROC (with the passed in values in tact) to handle any messages we ignored.
WndProcMain = CallWindowProc(lngPrevWndProc, hWnd, message, wParam, lParam)

End Function

Public Function ScrollText(strText As String) As String
'-------------------------------------------------------------------------------------
'This function scrolls text.
'-------------------------------------------------------------------------------------

'Take the first char of a string and move it to the back.
strText = (Right$(strText, Len(strText) - 1)) & Left$(strText, 1)
ScrollText = strText

End Function

Public Sub Main()

frmMain.Show

End Sub

Public Function ZoomForm(zoomToWhere As ZoomTypes, hWnd As Long) As Boolean
'-------------------------------------------------------------------------------------
'This function 'zooms' a window.
'-------------------------------------------------------------------------------------
Dim rctFrom As RECT
Dim rctTo As RECT
Dim lngTrayHand As Long
Dim lngStartMenuHand As Long
Dim lngChildHand As Long
Dim strClass As String * 255
Dim lngClassNameLen As Long
Dim lngRetVal As Long

'Select the type of zoom to do.
Select Case zoomToWhere

    'Zoom the window into the tray.
    Case ZOOM_FROM_TRAY
        'Get the handle to the start menu.
        lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)

        'Get the handle to the first child window of the start menu.
        lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

        'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
        Do
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
    
            'If it is the tray then store the handle.
            If InStr(1, strClass, "TrayNotifyWnd") Then
                lngTrayHand = lngChildHand
                Exit Do
            End If
    
            'If we didn't find it, go to the next sibling.
            lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)

        Loop

        'Get the RECT of  our form.
        lngRetVal = GetWindowRect(hWnd, rctFrom)
        
        'Get the RECT of the Tray.
        lngRetVal = GetWindowRect(lngTrayHand, rctTo)
        
        'Zoom from the tray to where our form is.
        lngRetVal = DrawAnimatedRects(frmMain.hWnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)

    Case ZOOM_TO_TRAY
        
        'Get the handle to the start menu.
        lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)

        'Get the handle to the first child window of the start menu.
        lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

        'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
        Do
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
    
            'If it is the tray then store the handle.
            If InStr(1, strClass, "TrayNotifyWnd") Then
                lngTrayHand = lngChildHand
                Exit Do
            End If
    
            'If we didn't find it, go to the next sibling.
            lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)

        Loop

        'Get the RECT of  our form.
        lngRetVal = GetWindowRect(hWnd, rctFrom)
        
        'Get the RECT of the Tray.
        lngRetVal = GetWindowRect(lngTrayHand, rctTo)
        
        'Zoom from where our form is to the tray .
        lngRetVal = DrawAnimatedRects(frmMain.hWnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)
            

End Select

End Function
