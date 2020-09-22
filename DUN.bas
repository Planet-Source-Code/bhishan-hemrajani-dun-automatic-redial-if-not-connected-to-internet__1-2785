Attribute VB_Name = "Module3"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const MAX_STRING_LENGTH As Integer = 256


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    ' RegQueryValueEx: If you declare the lpData parameter as String,
    '     you must pass it By Value.


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    ' Remote Access Services (RAS) APIs.
    Public Const RAS_MAXENTRYNAME As Integer = 256
    Public Const RAS_MAXDEVICETYPE As Integer = 16
    Public Const RAS_MAXDEVICENAME As Integer = 128
    Public Const RAS_RASCONNSIZE As Integer = 412


Public Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    End Type


Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
    End Type
    
    Public gstrISPName As String
    Public ReturnCode As Long


Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
