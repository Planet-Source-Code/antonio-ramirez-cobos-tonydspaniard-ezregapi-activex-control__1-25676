VERSION 5.00
Begin VB.UserControl EzRegApi 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   930
   DrawWidth       =   2
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   1080
   ScaleWidth      =   930
   ToolboxBitmap   =   "EzRegApi.ctx":0000
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   60
      Picture         =   "EzRegApi.ctx":0312
      Top             =   60
      Width           =   810
   End
End
Attribute VB_Name = "EzRegApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************'
'------------------------------------------------------'
' Project: EzRegAPI v1.0.23
'
' Date: July-28-2001
'
' Programmer: Antonio Ramirez Cobos
'
' Module: EzRegAPI ActiveX Control
'
' Description: Test application for EzRegAPI control. The control
'              is tested for most of the control's functionality but for
'              the methods that allows Association/removal of file types.
'              Note that this feature only works on Windows 9X platforms,
'              but it could easily be changed to work on NT or 2000
'              [like to be challenged?]
'
'              Please be careful with the examples of the included help
'              file EzRegAPI.chm, as the help file contains errata on some of its
'              examples [better stick with the examples shown on this
'              application]. The control was upgraded not long ago and
'              I didn't had the time to update its help file too.
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
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Enumerations
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Enum VALUE_TYPE
   regNone
   regSz  ' String
   regExpand_sz
   regBinary
   regDword
   regDwordLittleEndian = 4
   regDwordBigEndian
   regLink
   regMultiSz
   regResourceList
End Enum
Public Enum OPTIONS_QUERY_VALUE
   regValType
   regValdata
   regSizeValData
End Enum
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private moqvType As OPTIONS_QUERY_VALUE
Private mvtType As VALUE_TYPE
Private m_UseEvent As Boolean
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Event Error(ByVal ErrNumber As Long, ByVal ErrDescryption As String)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' About Sub procedure
'       Shows my fantastic about dialog box
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub About()
Attribute About.VB_UserMemId = -552
    Dim f As New frmAbout
    f.Show vbModal
    Set f = Nothing
End Sub
'********************************************************
' SetValue
'
' Purpose: Sets a value
' Inputs:
'       lKey: A handle to a Registry key [either HKEY_* or
'             from a previous call].
'       strValueName: The name of the value to be set.
'       ktDataType: The type of data stored in the value.
'       varDataValue: The value data packed in any form of
'                     valid VB data type. This will be converted
'                     to a 32-bit string.
'********************************************************
Public Sub SetValue(ByVal lKey As Long, _
                    ByVal strValueName As String, ByVal vtDataType As VALUE_TYPE, _
                    ByVal varDataValue As Variant)
Dim lReturn As Long ' Will be set with the value returned by the function

If vtDataType <> regDword Then
    ' Is not a DWORD value, we must use RegSetValueEx
    ' because the function is declared to handle string
    ' values. We use also this function to create binary
    ' values as well.
    ' Set value
    lReturn = RegSetValueEx(lKey, strValueName, 0&, vtDataType, CStr(varDataValue), Len(varDataValue) + 1)
Else
    ' DWORD value, call RegSetValueExA, as the function is
    ' declare it to handle DWORD values (numeric values!)
    lReturn = RegSetValueExA(lKey, strValueName, 0&, CInt(vtDataType), varDataValue, 4)
End If

If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
    ' Find out what happened and display error message
    Call ErrMsg(ERROR_SET_VALUE, "set a value", lReturn)
End If
End Sub ' SetValue


'********************************************************
' QueryValue
'
' Purpose: Lets you look up value data using the name
'          of the value stored in an open Registry key.
' Input:
'       lKeyNew: A handle to an open Registry key.
'       strValueName: The name of the value whose data
'                     is to be retrieved.
'       oqvType: Optional. Sets the type of information to
'                return. If not used the default value information
'                will be the data stored in the value.
'
' Returns:
'       1._ If oqvTYpe is not set or set to oqvType.ValueData,
'           then the function will return the data
'           stored in the value
'       2._ If set to oqvType.ValueType,
'           the function will return the type of value (i.e.,
'           REG_DWORD ["4"], REG_SZ ["1"], and so on)
'       3._ If set to oqvType.SizeValueData, the function will
'           return the size (in bytes) of the data
'
'      ''''''''''EVERYTHING IS RETURNED AS A STRING''''''''''''
'********************************************************
Public Function QueryValue(ByVal lKey As Long, ByVal strValueName As String, Optional ByVal oqvType As OPTIONS_QUERY_VALUE = 1) As String

Dim lReturn As Long        ' Will be set with the value returned by the function
Dim strValueData As String ' Will be set with the data of the Registry key
Dim stempVal As String      ' Temporary variable for use in type conversions
Dim ltempVal As Long       ' Temporary variable for use if type of value is DWORD
Dim lValueDataSz As Long   ' Will be set with the size [in bytes] of the data stored
Dim lType As Long          ' Will be set to indicate what type of data is stored
' Query Registry key
lReturn = RegQueryValueEx(lKey, strValueName, 0&, lType, ByVal 0&, lValueDataSz)
' No problems with the call?
If lReturn = ERROR_SUCCESS Then
    If lType = REG_DWORD Then
        'retrieve the key's content
        lReturn = RegQueryValueEx(lKey, strValueName, 0, 0, ltempVal, lValueDataSz)
    ElseIf lType = REG_SZ Or lType = REG_BINARY Then
    ' We also get from the binary values the 'string side' of it
        ' Create a buffer
        stempVal = String(lValueDataSz, Chr$(0))
        'retrieve the key's value
        lReturn = RegQueryValueEx(lKey, strValueName, 0, 0, ByVal stempVal, lValueDataSz)
        If lReturn = 0 Then
            'Remove the unnecessary chr$(0)'s
            stempVal = Left$(stempVal, InStr(1, stempVal, Chr$(0)) - 1)
        End If
    End If
End If
If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
   ' Find out what happened and display error message
   Call ErrMsg(ERROR_QUERY_VALUE, "query a value", lReturn)
   QueryValue = CStr(0) ' Return empty string
   
' No query options selected or options set to data of the
' value. Return the data of the value as string.
ElseIf oqvType = 1 Then
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lType   ' Search Data Types...
    Case REG_SZ, REG_BINARY      ' String Registry Key Data Type
        strValueData = stempVal      ' Return String Value
    Case REG_DWORD     ' Double Word Registry Key Data Type
        strValueData = CStr(ltempVal) ' Convert Double Word To String
    End Select
    ' Return result as a string
    QueryValue = Trim(strValueData)
' Options set to the type of value. Return type converted
' to string.
ElseIf oqvType = 0 Then
   QueryValue = CStr(lType)
Else ' Options set to the size of the data [in bytes]. Return its value
    ' converted to string.
   QueryValue = CStr(lValueDataSz)
End If
End Function ' QueryValue


'********************************************************
' EnumValueNames
'
' Purpose: Enumerates the names of all of the values
'          contained in an open Registry key.
' Input:
'       lKeyNew: A handle to an open Registry key.
'
' Returns:
'       1._ An empty string if not successful + Display Error msg
'       2._ Empty string if there is no more
'           values in the open key
'       3._ The name of the value at lIndex position if
'           successful
'
'********************************************************
Public Function EnumValueNames(ByVal lKey As Long) As String

Dim lReturn As Long         ' Will be set with the value returned by the function
Dim lType As Long           ' Will be set with the type of data stored in the value
Dim strNameValue As String  ' Will be set with to the name of the value
Dim lNameSz As Long         ' Specifies the buffer size to be allocated for strNameValue

' Allocate space
strNameValue = String(255, 0)
' String Size
lNameSz = 255

Static lIndex As Long  ' Keeps track of the sequence number
Static lKeyOld As Long ' Stores the last handle to the Registry key

' First time?
If IsEmpty(lIndex) Then lIndex = 0
' Check if it is a new call. If True then
' set lKeyOld to the new key, lIndex to 0
' to be able to enumerate the subkeys of the new key.
If lKeyOld <> lKey Then ' If the handle is different
   lKeyOld = lKey ' Set lKeyOld to the new handle
   lIndex = 0 ' Initiate the sequence
End If
' Value at lIndex position
lReturn = RegEnumValue(lKeyOld, lIndex, strNameValue, 255, 0&, ByVal 0&, ByVal 0, ByVal 0&)
If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
    If lReturn = ERROR_NO_MORE_ITEMS Then ' No more values?
        lIndex = 0 ' Restart Enumeration
        EnumValueNames = NO_MORE_ITEMS ' Return empty string
    Else ' An error has occurred
        ' Find out what happened and display error message
        Call ErrMsg(ERROR_ENUM_VALUES, "enumerate value names", lReturn)
        lIndex = 0 ' Re-set counter
        EnumValueNames = "" ' Return an empty string
    End If
Else ' Successful, return the name of the value
   EnumValueNames = Left(strNameValue, lNameSz) 'InStr(strNameValue, Chr(0)) - 1)
   ' Add one to sequence
   lIndex = lIndex + 1
End If
End Function ' EnumValueNames

'********************************************************
' EnumKeyNames
'
' Purpose: Enumerates the names of all of the subkeys
'          directly under an open Registry key.
' Input:
'       lKeyNew: A handle to an open Registry key.
'
' Returns:
'       1._ An empty string if not successful + display error MSG
'       2._ Empty string if there is no more
'           values in the open key
'       3._ The name of the subkey at lIndex position if
'           successful
'
'********************************************************
Public Function EnumKeyNames(ByVal lKey As Long) As String

Dim lReturn As Long     ' Will be set with the value returned by the function
Dim strName As String   ' Will be set with the name of key
Dim lNameSz As Long
Dim fFiletime As FILETIME
Static lIndex As Long   ' Keeps track of the sequence number'
Static lKeyOld As Long  ' Stores the last handle to the Registry key

' Allocate space
strName = String(255, 0)
' String size
lNameSz = 255

' First time?
If IsEmpty(lIndex) Then lIndex = 0
' Check if it is a new call. If True then
' set lKeyOld to the new key, lIndex to 0
' to be able to enumerate the subkeys of the new key.
If lKeyOld <> lKey Then
   lKeyOld = lKey ' Store new handle to lKeyOld
   lIndex = 0   ' Initiate sequence
End If
' Key name at lIndex position
lReturn = RegEnumKeyEx(lKeyOld, lIndex, strName, lNameSz, 0, vbNullString, ByVal 0&, fFiletime)
        
If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
    If lReturn = ERROR_NO_MORE_ITEMS Then ' No more keys?
        lIndex = 0 ' Restart Enumeration
        EnumKeyNames = NO_MORE_ITEMS ' Return empty string
    Else ' An error has occurred
        ' Find out what happened and display error message
        Call ErrMsg(ERROR_ENUM_KEYS, "enumerate subkeys", lReturn)
        lIndex = 0 ' Re-set counter
        EnumKeyNames = ""  ' Return empty string
    End If
Else ' Success, return the name of the key
   EnumKeyNames = Left(strName, InStr(strName, Chr(0)) - 1)
   ' Add one to sequence
   lIndex = lIndex + 1
End If
End Function ' EnumKeyNames

'********************************************************
' DeleteValue
'
' Purpose: Deletes a value from an open Registry key.
' Input:
'       lKey: A handle to an open Registry key.
'       strValueName: The name of the value to delete.
'
'********************************************************
Public Sub DeleteValue(ByVal lKey As Long, ByVal strValueName As String)

Dim lReturn As Long ' Will be set with the value returned by the function
 ' Delete the value
lReturn = RegDeleteValue(lKey, strValueName)
 
If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
    ' Find out what happened and display error message
    Call ErrMsg(ERROR_DELETE_VALUE, "delete a value", lReturn)
End If

End Sub      ' DeleteValue

'********************************************************
' DeleteKey
'
' Purpose: Deletes a subkey of an open Registry key provided
'          that the subkey contains no subkeys of its own [but
'          the subkey may contain values].
' Input:
'       lKey: A handle to an open Registry key.
'       strSubKey: The name of the subkey to delete
'********************************************************
Public Sub DeleteKey(ByVal lKey As Long, ByVal strSubKey As String)

Dim lReturn As Long ' Will be set with the value returned by the function
' Delete subkey
lReturn = RegDeleteKey(lKey, strSubKey)

If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
  ' Find out what happened and display error message
  Call ErrMsg(ERROR_DELETE_KEY, "delete a key", lReturn)
End If
End Sub      ' Deletekey

'********************************************************
' CreateKey
'
' Purpose: Creates a new Registry subkey.
' Input:
'       lKey: A handle to an open Registry key.
'       strKeyName: The name of the new subkey to be created
'       lAccess: A numeric mask of bits specifying what type
'                of access is desired when opening the new subkey:
'                  a/ RegistryAPI.ReadAccess
'                  b/ RegistryAPI.WriteAccess
'                  c/ RegistryAPI.TotalAccess
' Returns:
'       1._ Raise error if not successful
'       2._ The handle to the new created subkey
'
'********************************************************
Public Function CreateKey(ByVal lKey As Long, ByVal strKeyName As String, ByVal lAccess As Long) As Long

Dim lReturn As Long ' Will be set with the value returned by the function
Dim lHandle As Long ' Will be set with the handle of the new created subkey
Dim lRet As Long
CreateKey = FAILURE
' Create subkey
lReturn = RegCreateKeyEx(lKey, strKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, _
          lAccess, ByVal 0&, lHandle, lRet)
          
If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
   ' Find out what happened and display error message
   Call ErrMsg(ERROR_CREATE_KEY, "create a key", lReturn)
' Success, return the handle of the new created subkey
Else: CreateKey = lHandle
End If
End Function ' CreateKey

'********************************************************
' ConnectRegistry
'
' Purpose: Connects to one of the root Registry keys of a
'          remote computer.
' Input:
'       strComputerName: Is the name [or address] of a remote
'                        computer whose Registry you wish to
'                        access (i.e., \\computername).
'       lRootKey: Must be either HKEY_LOCAL_MACHINE or HKEY_USERS
'                 and specifies which root Registry key on the
'                 remote computer you wish to access.
'
' Returns:
'       1._ Zero if not successful
'       2._ The handle of the remote registry if successful
'
'********************************************************
Public Function ConnectRegistry(ByVal strComputerName As String, ByVal lRootKey As Long) As Long

Dim lReturn As Long ' Will be set with the value returned by the function
Dim lHandle As Long ' Will be set with the handle of the remote registry key
ConnectRegistry = FAILURE
' Connect to remote registry
lReturn = RegConnectRegistry(strComputerName, lKey, lHandle)

If lReturn <> ERROR_SUCCESS Then  ' Unsuccessful?
  ' Find out what happened and display error message
  Call ErrMsg(ERROR_CONNECT_REG, "connect to a remote registry", lReturn)
End If
  ConnectRegistry = lHandle
End Function ' ConnectRegistry

'********************************************************
' CloseKey
'
' Purpose: Closes the handle to a Registry key.
' Input:
'       lKey: A handle to an opened Registry key
'
'********************************************************
Public Sub CloseKey(ByVal lKey As Long)

Dim lReturn As Long
' Close registry key
lReturn = RegCloseKey(lKey)
If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
    ' Find out what happened and display error message
    Call ErrMsg(ERROR_CLOSE_KEY, "close a key", lReturn)
End If
End Sub      ' CloseKey

'********************************************************
' Openkey
'
' Purpose: Opens an existing Registry key.
' Input:
'       lKey: Is the handle to a Registry key [ either HKEY_* or
'             from a previous call].
'       strSubKey: Is the name of an existing subkey to be opened.
'                  Can be "" to open an additional handle to the
'                  key specified by lKey.
'       lAccess: A numeric mask of bits specifying what type
'                of access is desired when opening the new subkey:
'                  a/ RegistryAPI.ReadAccess
'                  b/ RegistryAPI.WriteAccess
'                  c/ RegistryAPI.TotalAccess
' Returns:
'       1._ Zero if not successful [Error msg]
'       2._ The handle of the Registry key if successful
'
'********************************************************
Public Function OpenKey(ByVal lKey As Long, ByVal strSubKey As String, ByVal lAccess As Long) As Long

Dim lHandle As Long    'Handle to the opened key if the function succeeds
Dim lReturn As Long    'Checks the value returned by the function

OpenKey = FAILURE ' Pessimistic
' Open Registry key
lReturn = RegOpenKeyEx(lKey, strSubKey, 0&, lAccess, lHandle)

If lReturn <> ERROR_SUCCESS Then ' Unsuccessful?
   ' Display error message!
   ' Find out with FormatMessage function which error was
   ' and display it.
   Call ErrMsg(ERROR_OPEN_KEY, "open a key", lReturn)
End If
OpenKey = lHandle ' Success!, return the handle of the opened key
End Function ' Openkey

'********************************************************
' AssociateExtensions
'
' Purpose: Associates specified file extensions to the windows
'          explorer, so the programmer can personalize its file types
'          and users can double-click the files to open the program.
'          Obviously, the program must check for the Command at the
'          application's load event and act accordingly.
' Input:
'       AppTitle: The title of the application. Actually, it is the
'                registry key to which is related the file extension.
'                If you are registering more than one file type, then
'                use two different names on this parameter
'       AppPath: Is the application's path. Windows will follow this
'                path to find the program that handles the file type
'       AppEXEName: The program used to handle the file type
'       FileExtension: The extension to register [i.e.: .frx]
'       FileType: A description of the file type [make it small as this appears
'                 on the description column of explorer
'       IconFileName: Path to Icon related to the file type [explorer will display it]
'                     It can be the path to an icon file, the path to the
'                     dll were to extract the icon from [i.e. C:\mydll.dll,1],
'                     or the path to the EXE were to extract the icon from.
'       Parameters: Extra parameters attached to the Command.
'                   I.e.: "C:\mydir\myexe.exe" "%1 /b/c"
'                   extra parameters will be /b/c and these will be attach
'                   to the Command string when the file opens the application.
'                   For example, if a file with my file extensions named "A.abc"
'                   on the C:\ root is double-clicked
'                   on explorer and the above example was the Associated command
'                   the Command passed to my application once open will be
'                   "C:\A.abc /b/c"
'********************************************************
Public Sub AssociateExtensions(AppTitle As String, AppPath As String, _
AppEXEName As String, FileExtension As String, FileType As String, _
IconFileName As String, Optional Parameters As String)
   If AppTitle = "" Then Exit Sub
   Dim sKeyName As String   ' Holds Key Name in registry.
   Dim sKeyValue As String  ' Holds Key Value in registry.
   Dim ret&           ' Holds error status if any from API calls.
   Dim lRet As Long ' Holds if the key was created new or an existing was opened
   Dim lphKey&        ' Holds  key handle from RegCreateKey.

   If Right(AppPath, 1) <> "\" Then
      AppPath = AppPath & "\"
   End If

' This creates a Root entry that called as the string of AppTitle.
   sKeyName = AppTitle
   sKeyValue = FileType
   ret& = RegCreateKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
   ret& = RegSetValueEx(lphKey&, "", 0&, REG_SZ, sKeyValue, 0&)
   ret& = RegCloseKey(lphKey&)
'   This creates a Root entry called as the string of FileExtension
'   associated with AppTitle.
   sKeyName = FileExtension
   sKeyValue = AppTitle
   ret& = RegCreateKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
   ret& = RegSetValueEx(lphKey&, "", 0&, REG_SZ, sKeyValue, 0&)
   ret& = RegCloseKey(lphKey&)
' This sets the command line for AppTitle.
   sKeyName = AppTitle
   If Parameters <> "" Then
        sKeyValue = """" & AppPath & AppEXEName & ".exe"" " & Trim(Parameters) & " ""%1"""
   Else
        sKeyValue = """" & AppPath & AppEXEName & ".exe"" %1"
   End If
   ret& = RegCreateKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
   sKeyName = "shell"
   ret& = RegCreateKeyEx(lphKey&, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
   sKeyName = "open"
   ret& = RegCreateKeyEx(lphKey&, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
   sKeyName = "command"
   ret& = RegCreateKeyEx(lphKey&, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
   ret& = RegSetValueEx(lphKey&, "", 0&, REG_SZ, sKeyValue, MAX_PATH)
   ret& = RegCloseKey(lphKey&)
' This sets the icon for the file extension
   If IconFileName <> "" Then
    sKeyName = AppTitle
    sKeyValue = IconFileName
    ret& = RegCreateKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
    ret& = RegCreateKeyEx(lphKey&, "DefaultIcon", 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lphKey&, lRet)
    ret& = RegSetValueEx(lphKey&, "", 0&, REG_SZ, sKeyValue, MAX_PATH)
    ret& = RegCloseKey&(lphKey&)
   End If
 
' This notifies the shell that the icon has changed
  SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
 
End Sub
'********************************************************
' AddCommandToExtensionsAssoc
'
' Purpose: Adds an extra command to the file associations. This
'          extra command will appear on the context menu displayed
'          by explorer when the user right-clicks your defined file type
' Input:
'       AppTitle: The registry key to which is related the file extension.
'                If you are registering more than one file type, then
'                use two different names on this parameter
'       AppPath: Is the application's path. Windows will follow this
'                path to find the program that handles the file type
'       AppCommand: The command to be added
'       Parameters: Extra parameters attached to the Command.
'                   I.e.: "C:\mydir\myexe.exe" "%1 /b/c"
'                   extra parameters will be /b/c and these will be attach
'                   to the Command string when the file opens the application.
'                   For example, if a file with my file extensions named "A.abc"
'                   on the C:\ root is double-clicked
'                   on explorer and the above example was the Associated command
'                   the Command passed to my application once open will be
'                   "C:\A.abc /b/c"
'********************************************************
Public Sub AddCommandToExtensionsAssoc(AppTitle As String, AppPath As String, _
AppCommand As String, AppEXEName As String, Optional Parameters As String)
    Dim lhKey As Long ' Holds handle to key
    Dim sKeyValue As String ' Holds key value
    Dim sKeyName As String ' Holds key name
    Dim ret As Long
    
    If Right(AppPath, 1) <> "\" Then
      AppPath = AppPath & "\"
    End If
    If Parameters <> "" Then
        sKeyValue = """" & AppPath & AppEXEName & ".exe"" " & Trim(Parameters) & """%1"""
    Else
        sKeyValue = """" & AppPath & AppEXEName & ".exe"" %1"
    End If
    ' This adds the extra command to registered file type extensions
    sKeyName = AppTitle & "\shell\" & AppCommand
    ret = RegCreateKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lhKey, ByVal 0&)
    sKeyName = "command"
    ret = RegCreateKeyEx(lhKey, sKeyName, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lhKey, ByVal 0&)
    ret = RegSetValueEx(lhKey, "", 0&, REG_SZ, sKeyValue, MAX_PATH)
    ret = RegCloseKey(lhKey)
End Sub
'********************************************************
' RemoveExtensionsAssoc
'
' Purpose: Removes file association.
' Input:
'       AppTitle: The registry key to which is related the file extension.
'       FileExtension: Associated file extension
'********************************************************
Public Sub RemoveExtensionsAssoc(AppTitle As String, FileExtension As String)
    'Delete all keys
    RegDeleteKey HKEY_CLASSES_ROOT, FileExtension
    RegDeleteKey HKEY_CLASSES_ROOT, AppTitle
    'Notify shell on the delete and refresh the icons
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
'********************************************************
' BackupExtensionsAssoc
'
' Purpose: Creates a back up REGistry file for file registered
'          file extensions.
'********************************************************
Public Sub BackupExtensionsAssoc(FileName As String, AppTitle As String, FileExtension As String, Optional FileType As String)
On Error Resume Next
    Kill FileName
    Dim RegFile As String, Buffer As String, Val As Long, FileNumber As Integer
    RegFile = "REGEDIT4" & vbCrLf & vbCrLf
    RegFile = RegFile & "[HKEY_CLASSES_ROOT\" & FileExtension & "]"
    RegFile = RegFile & vbCrLf & "@=" & Chr(34) & AppTitle & Chr(34)
    If FileType <> "" Then
        RegFile = RegFile & vbCrLf & vbCrLf & "[HKEY_CLASSES_ROOT\" & AppTitle & "]"
        RegFile = RegFile & vbCrLf & "@=" & Chr(34) & FileType & Chr(34)
    End If
    RegFile = RegFile & vbCrLf & vbCrLf & "[HKEY_CLASSES_ROOT\" & AppTitle & "]"
    RegFile = RegFile & vbCrLf & vbCrLf & "[HKEY_CLASSES_ROOT\" & AppTitle & "\shell\" & "]"
    RegFile = RegFile & vbCrLf & vbCrLf & "[HKEY_CLASSES_ROOT\" & AppTitle & "\shell\open" & "]"
    RegFile = RegFile & vbCrLf & vbCrLf & "[HKEY_CLASSES_ROOT\" & AppTitle & "\shell\open\command" & "]"
    Buffer = QueryValue(HKEY_CLASSES_ROOT, AppTitle & "\shell\open\command")
    RegFile = RegFile & vbCrLf & "@=" & Chr(34) & Trim(Replace(Trim(Buffer), "\", "\\") & Chr(34))
    Buffer = QueryValue(HKEY_CLASSES_ROOT, AppTitle & "\DefaultIcon")
    RegFile = RegFile & vbCrLf & vbCrLf & "[HKEY_CLASSES_ROOT\" & AppTitle & "\DefaultIcon" & "]"
    RegFile = RegFile & vbCrLf & "@=" & Chr(34) & Trim(Replace(Trim(Buffer), "\", "\\") & Chr(34))
    FileNumber = FreeFile
    Open FileName For Binary As FileNumber Len = Len(RegFile)
        Put #FileNumber, , RegFile
    Close FileNumber
End Sub
'********************************************************
' RestoreAssoc
'
' Purpose: Merges created REGistry file to the registry
'********************************************************
Function RestoreAssoc(FileName As String)
    Shell "Regedit.exe /s" & FileName
End Function

'********************************************************
' START PROPERTIES
'********************************************************
' Returns a handle to HKEY_DYN_DATA registry key
Public Property Get HkeyDynData() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HkeyDynData
    HkeyDynData = HKEY_DYN_DATA
End Property
' Returns a handle to HKEY_CURRENT_USER registry key
Public Property Get HkeyCurrentUser() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HkeyCurrentUser
    HkeyCurrentUser = HKEY_CURRENT_USER
End Property
' Returns a handle to HKEY_CURRENT_CONFIG
Public Property Get HkeyCurrentConfig() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HkeyCurrentConfig
    HkeyCurrentConfig = HKEY_CURRENT_CONFIG
End Property

' Returns a handle to HKEY_CLASSESS_ROOT
Public Property Get HkeyClassesRoot() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hkeyClassesRoot
    HkeyClassesRoot = HKEY_CLASSES_ROOT
End Property
' Returns a handle to HKEY_LOCAL_MACHINE registry key
Public Property Get HkeyLocalMachine() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HkeyLocalMachine
    HkeyLocalMachine = HKEY_LOCAL_MACHINE
End Property

' Returns a handle to HKEY_USERS registry key
Public Property Get HkeyUsers() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HkeyUsers
    HkeyUsers = HKEY_USERS
End Property

' Returns the value of KEY_READ
Public Property Get ReadAccess() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReadAccess
    ReadAccess = KEY_READ
End Property

' Returns the value of KEY_WRITE
Public Property Get WriteAccess() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WriteAccess
    WriteAccess = KEY_WRITE
End Property

' Returns the value of KEY_ALL_ACCESS
Public Property Get TotalAccess() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalAccess
    TotalAccess = KEY_ALL_ACCESS
End Property

' Returns the value of KEY_SET_VALUE
Public Property Get SetValueAccess() As Long
    SetValueAccess = KEY_SET_VALUE
End Property

' Returns the value of KEY_QUERY_VALUE
Public Property Get QueryValueAccess() As Long
    QueryValueAccess = KEY_QUERY_VALUE
End Property

' Returns the value of KEY_ENUMERATE_SUB_KEYS
Public Property Get EnumSubKeysAccess() As Long
    EnumSubKeysAccess = KEY_ENUMERATE_SUB_KEYS
End Property

' Returns the value of KEY_CREATE_LINK
Public Property Get CreateLinkAccess() As Long
    CreateLinkAccess = KEY_CREATE_LINK
End Property

' Returns the value of KEY_CREATE_SUB_KEY
Public Property Get CreateSubKeyAccess() As Long
    CreateSubKeyAccess = KEY_CREATE_SUB_KEY
End Property

' This property is used to specify which type of
' value is to be set with the function SetKeyValue.
Public Property Get ValueType() As VALUE_TYPE
'used when retrieving value of a property, on the right side of an assignment.
  ValueType = mvtType ' Return the type of key to set or query
End Property


Public Property Let ValueType(ByVal NewVType As VALUE_TYPE)
'used when setting the value of a property
'KType is used to set the type of data stored in a value
  Select Case NewVType
     Case 0 To 8
         ' Setting is valid.
           mvtType = NewVType
     Case Else
         ' Setting non-valid. Display error message.
         Err.Raise Number:=vbObjectError + 32112, _
         Description:="Invalid ValueType setting."
  End Select
  
End Property

' This property is used to set the parameters of QueryValue
' function. You can specify what type of query would you like
' to do to the value: TYPE, VALUE, or SIZE OF DATA.
Public Property Get OptQueryValue() As OPTIONS_QUERY_VALUE
'used when retrieving value of a property
OptQueryValue = moqvType

End Property

Public Property Let OptQueryValue(ByVal NewOption As OPTIONS_QUERY_VALUE)

   Select Case NewOption
   Case 0 To 2
        ' Setting is valid
        moqvType = NewOption
   Case Else
        ' Setting is non-valid. Display error message.
        Err.Raise Number:=vbObjectError + 32113, _
        Description:="Invalid OptQueryValue setting."
   End Select
   
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_UseEvent = PropBag.ReadProperty("UseErrorEvent", False)
End Sub

'********************************************************
' END PROPERTIES
'********************************************************
Private Sub UserControl_Resize()
 ' Don't allow changes to width and height

    Dim R As RECT
    Size DEF_WIDTH, DEF_HEIGHT
    UserControl.ScaleMode = vbPixels
    SetRect R, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    DrawEdge hdc, R, EDGE_RAISED, BF_ADJUST Or BF_RECT
End Sub


Private Sub ErrMsg(ByVal lErrorConst As Long, ByVal strOperation As String, ByVal _
                              lErrorMsg As Long)
Dim lMsgReturn As Long
Dim strMessage As String * 1024
Dim lMsgSz As Long
lMsgSz = 1024&

' Find out what is the error message from the system
' and raise the event
lMsgReturn = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, SUBLANG_DEFAULT, lErrorMsg, 0&, strMessage, lMsgSz, 0)
If lMsgReturn = ERROR_MSG_FAIL Then ' FormatMessage function has failed.
' Raise the Error Event with a generic error message
    If m_UseEvent Then
        RaiseEvent Error(vbObjectError + lErrorConst, "An error has occurred when trying to " & strOperation & "." & vbCrLf & _
              "EzRegApi couldn't build an appropriate error message" & vbCrLf & _
              "due to a failure of the 'FormatMessage' function in your system.")
    Else
        Err.Raise vbObjectError + lErrorConst, "EzRegApi", "An error has occurred when trying to " & strOperation & "." & vbCrLf & _
              "EzRegApi couldn't build an appropriate error message" & vbCrLf & _
              "due to a failure of the 'FormatMessage' function in your system."
    End If
Else ' Display error
    If m_UseEvent Then
            RaiseEvent Error(vbObjectError + lErrorConst, "An error has occurred when trying to " & strOperation & "." & vbCrLf & _
                "Error: " & vbCrLf & Left(strMessage, lMsgReturn))
    Else
           Err.Raise vbObjectError + lErrorMsg, "EzRegApi", "An error has occurred when trying to " & strOperation & "." & vbCrLf & _
                "Error: " & Left(strMessage, lMsgReturn)
    End If
End If
End Sub
Public Property Get UseErrorEvent() As Boolean
    UseErrorEvent = m_UseEvent
End Property

Public Property Let UseErrorEvent(ByVal vNewValue As Boolean)
    m_UseEvent = vNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "UseErrorEvent", m_UseEvent, False
End Sub



