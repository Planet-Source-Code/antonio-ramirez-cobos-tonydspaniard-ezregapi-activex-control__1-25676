Attribute VB_Name = "Module1"

'******************************************************'
'------------------------------------------------------'
' Project: EzRegAPI v1.0.23
'
' Date: July-28-2001
'
' Programmer: Antonio Ramirez Cobos
'
' Module: Module1
'
' Description: API Stuff
'
'              From the Author:
'              'cause I consider myself in a continuous learning
'              path with no end on programming, please, if you
'              can improve this program
'              contact me at: *TONYDSPANIARD@HOTMAIL.COM*
'
'              I would be pleased to hear from your opinions,
'              suggestions, and/or recommendations. Also, if you
'              know something I don't know and wish to share it
'              with me, here you'll have your techy pal from Spain
'              that will do exactly the same towards you. If I can
'              help you in any way, just ask.
'
'              INTELLECTUAL COPYRIGHT STUFF [Is up to you anyway]
'              --------------------------------------------------
'              This code is copyright 2001 Antonio Ramirez Cobos
'              This code may be reused and modified for non-commercial
'              purposes only as long as credit is given to the author
'              in the programmes about box and it's documentation.
'              If you use this code, please email me at:
'              TonyDSpaniard@hotmail.com and let me know what you think
'              and what you are doing with it.
'
'              PS: Don't forget to vote for me buddy programmer!
'------------------------------------------------------'
'******************************************************'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Registry Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long

Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, lpData As Any, ByRef lpcbData As Long) As Long      ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long                                 ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Draw Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Draw Constants
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
'Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const STANDARD_RIGHTS_ALL& = &H1F0000
Public Const READ_CONTROL& = &H20000
Public Const KEY_QUERY_VALUE& = &H1
Public Const KEY_CREATE_SUB_KEY& = &H4
Public Const KEY_ENUMERATE_SUB_KEYS& = &H8
Public Const KEY_CREATE_LINK& = &H20
Public Const KEY_SET_VALUE& = &H2
Public Const SYNCHRONIZE& = &H100000
Public Const KEY_NOTIFY& = &H10
Public Const KEY_WRITE& = ((READ_CONTROL Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_READ& = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS& = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const HKEY_CLASSES_ROOT& = &H80000000
Public Const HKEY_CURRENT_CONFIG& = &H80000005
Public Const HKEY_CURRENT_USER& = &H80000001
Public Const HKEY_DYN_DATA& = &H80000006
Public Const HKEY_LOCAL_MACHINE& = &H80000002
Public Const HKEY_USERS& = &H80000003
Public Const ERROR_SUCCESS& = 0&
Public Const REG_OPTION_NON_VOLATILE = 0&         ' Key is preserved when system is rebooted
Public Const REG_SZ& = 1&                         ' Unicode nul terminated string
Public Const REG_NONE& = 0&                       ' No value type
Public Const REG_MULTI_SZ& = 7&                   ' Multiple Unicode strings
Public Const REG_LINK& = 6&                       ' Symbolic Link (unicode)
Public Const REG_EXPAND_SZ& = 2&                  ' Unicode nul terminated string
Public Const REG_DWORD_LITTLE_ENDIAN& = 4&        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN& = 5&           ' 32-bit number
Public Const REG_DWORD& = 4&                      ' 32-bit number
Public Const REG_BINARY& = 3&                     ' Free form binary
Public Const REG_RESOURCE_LIST& = 8               ' Resource list in the resource map
Public Const NO_MORE_ITEMS As String = ""
Public Const ERROR_NO_MORE_ITEMS& = 259&
Public Const FAILURE& = 0
Public Const SUCCESS& = 1
Public Const FORMAT_MESSAGE_FROM_SYSTEM& = &H1000
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER& = &H100
Public Const LANG_NEUTRAL& = &H0
Public Const SUBLANG_DEFAULT& = &H1 '  user default
Public Const MAXMIN_WIDTH& = 540&
Public Const MAXMIN_HEIGHT& = 500
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Error Constants
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const ERROR_SET_VALUE = 1000&
Public Const ERROR_QUERY_VALUE = 1001&
Public Const ERROR_DELETE_VALUE = 1002&
Public Const ERROR_CREATE_KEY = 1003&
Public Const ERROR_DELETE_KEY = 1004&
Public Const ERROR_ENUM_KEYS = 1005&
Public Const ERROR_ENUM_VALUES = 1006&
Public Const ERROR_CLOSE_KEY = 1007&
Public Const ERROR_OPEN_KEY = 1008&
Public Const ERROR_CONNECT_REG = 1009&
Public Const ERROR_MSG_FAIL& = &H0
' Default size
Public Const DEF_WIDTH = 930&
Public Const DEF_HEIGHT = 1080&

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Explorer
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Declare Sub SHChangeNotify Lib "shell32.dll" _
           (ByVal wEventId As Long, _
            ByVal uFlags As Long, _
            dwItem1 As Any, _
            dwItem2 As Any)

Const SHCNE_ASSOCCHANGED = &H8000000
Const SHCNF_IDLIST = &H0&
Public Const MAX_PATH = 260



