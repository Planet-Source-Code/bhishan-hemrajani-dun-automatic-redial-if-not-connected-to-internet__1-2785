VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "247"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtConnection 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Mom's Internet"
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdMinTray 
      Caption         =   "Minimize to System Tray"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   240
   End
   Begin VB.CheckBox chckAuto 
      Caption         =   "Automatically reconnect to internet"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1920
      Top             =   240
   End
   Begin VB.Timer tmrUptime 
      Interval        =   1
      Left            =   1440
      Top             =   240
   End
   Begin VB.Timer LoopsOnce 
      Interval        =   1
      Left            =   960
      Top             =   240
   End
   Begin VB.Label lblLoop 
      Alignment       =   2  'Center
      Caption         =   "You have been reconnected 0 times."
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblUptime 
      Alignment       =   2  'Center
      Caption         =   "Uptime"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblUptimeTitle 
      Alignment       =   2  'Center
      Caption         =   "Windows Uptime:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblConnection 
      Alignment       =   2  'Center
      Caption         =   "Name of Connection to use:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Tray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Window"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loops As Integer
Private Sub chckAuto_Click()
If chckAuto.Value = Checked Then
    tmrReconnect.Enabled = True
    Dim lngRetVal As Long

    'Get the hwnd from of the form.
    lngHwnd = frmMain.hWnd

    'ID used for callback.
    lngWndID = App.hInstance

    'Initialize the icon tip.
    strToolTip = "Double Click Here to Make 247 or Right click for a menu....." & Chr(0)

    'Subclass the window.
    lngPrevWndProc = SetWindowLong(lngHwnd, GWL_WNDPROC, AddressOf WndProcMain)

    'Set all Tray Icon info.
    nidTray.hWnd = lngHwnd
    nidTray.cbSize = Len(nidTray)
    nidTray.hIcon = frmMain.Icon
    nidTray.szTip = strToolTip
    nidTray.uCallbackMessage = WM_CALLBACK_MSG
    nidTray.uID = lngWndID
    nidTray.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON

    'Call windows to add the icon to the tray.
    lngRetVal = Shell_NotifyIconA(NIM_ADD, nidTray)

    ZoomForm ZOOM_TO_TRAY, Me.hWnd

    'Hide our form.
    frmMain.Hide
Else
    tmrReconnect.Enabled = False
    End If
End Sub

Private Sub cmdMinTray_Click(Index As Integer)
Dim lngRetVal As Long

'Get the hwnd from of the form.
lngHwnd = frmMain.hWnd

'ID used for callback.
lngWndID = App.hInstance

'Initialize the icon tip.
strToolTip = "Double Click Here to Make 247 Appear or Right click for a menu....." & Chr(0)

'Subclass the window.
lngPrevWndProc = SetWindowLong(lngHwnd, GWL_WNDPROC, AddressOf WndProcMain)

'Set all Tray Icon info.
nidTray.hWnd = lngHwnd
nidTray.cbSize = Len(nidTray)
nidTray.hIcon = frmMain.Icon
nidTray.szTip = strToolTip
nidTray.uCallbackMessage = WM_CALLBACK_MSG
nidTray.uID = lngWndID
nidTray.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON

'Call windows to add the icon to the tray.
lngRetVal = Shell_NotifyIconA(NIM_ADD, nidTray)

ZoomForm ZOOM_TO_TRAY, Me.hWnd

'Hide our form.
frmMain.Hide
End Sub

Private Sub Form_Load()
chckAuto.Value = Checked
End Sub

Private Sub LoopsOnce_Timer()
If Loops = 1 Then
    lblLoop.Caption = "You have been reconnected once."
End If
End Sub

Private Sub mnuExit_Click()
SetWindowLong lngHwnd, GWL_WNDPROC, lngPrevWndProc
Shell_NotifyIconA NIM_DELETE, nidTray
Call CloseApp
End Sub
Public Function Connected_To_ISP() As Boolean


    Dim hKey As Long
    Dim lpSubKey As String
    Dim phkResult As Long
    Dim lpValueName As String
    Dim lpReserved As Long
    Dim lpType As Long
    Dim lpData As Long
    Dim lpcbData As Long
    Connected_To_ISP = False
    lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
    ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)


    If ReturnCode = ERROR_SUCCESS Then
        hKey = phkResult
        lpValueName = "Remote Connection"
        lpReserved = APINULL
        lpType = APINULL
        lpData = APINULL
        lpcbData = APINULL
        ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
        lpcbData = Len(lpData)
        ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData)


        If ReturnCode = ERROR_SUCCESS Then


            If lpData = 0 Then
                ' Not Connected
            Else
                ' Connected
                Connected_To_ISP = True
            End If

        End If

        RegCloseKey (hKey)
    End If

End Function


Public Sub CloseApp()

Dim f As Form

For Each f In Forms
    Unload f
Next

End Sub

Private Sub mnuRestore_Click()
SetWindowLong lngHwnd, GWL_WNDPROC, lngPrevWndProc

'Delete our icon from the tray.
Shell_NotifyIconA NIM_DELETE, nidTray
frmMain.Show
End Sub

Private Sub Timer1_Timer()
Dim isp As Integer
isp = FindWindow("#32770", "Connecting to " & txtConnection.Text)
If isp <> 0 Then
    tmrReconnect.Enabled = False
Else
    If chckAuto.Value = Checked Then
        tmrReconnect.Enabled = True
    End If
End If
End Sub





Private Sub tmrReconnect_Timer()
    If Connected_To_ISP = False Then
    Dim res
    res = Shell("rundll32.exe rnaui.dll,RnaDial " & txtConnection.Text, 1)
    DoEvents
    SendKeys "{enter}", True
    DoEvents
    Loops = Loops + 1
    lblLoop.Caption = "You have been reconnected " & Loops & " times."
End If
End Sub

Private Sub tmrUptime_Timer()
Dim Secs, Mins, Hours, Days
    Dim TotalMins, TotalHours, TotalSecs, TempSecs
    Dim CaptionText
    TotalSecs = Int(GetTickCount / 1000)
    Days = Int(((TotalSecs / 60) / 60) / 24)
    TempSecs = Int(Days * 86400)
    TotalSecs = TotalSecs - TempSecs
    TotalHours = Int((TotalSecs / 60) / 60)
    TempSecs = Int(TotalHours * 3600)
    TotalSecs = TotalSecs - TempSecs
    TotalMins = Int(TotalSecs / 60)
    TempSecs = Int(TotalMins * 60)
    TotalSecs = (TotalSecs - TempSecs)


    If TotalHours > 23 Then
        Hours = (TotalHours - 23)
    Else
        Hours = TotalHours
    End If



    If TotalMins > 59 Then
        Mins = (TotalMins - (Hours * 60))
    Else
        Mins = TotalMins
    End If

    CaptionText = Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " seconds" & vbCrLf
    'CaptionText = CaptionText & "TickCount: " & TotalSecs & vbCrLf
    lblUptime.Caption = CaptionText
End Sub
