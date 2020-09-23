VERSION 5.00
Object = "{1E4B17B2-2F6F-11D4-BFBD-D9FFEE979A03}#14.0#0"; "EzRegApi.ocx"
Begin VB.Form TestForm 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin RegApi.EzRegApi EzRegApi1 
      Left            =   4005
      Top             =   495
      _ExtentX        =   1640
      _ExtentY        =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   396
      Left            =   4050
      TabIndex        =   10
      Top             =   4305
      Width           =   990
   End
   Begin VB.CommandButton CmdCreateKey 
      Caption         =   "Create &Key"
      Height          =   396
      Left            =   165
      TabIndex        =   9
      Top             =   3382
      Width           =   1896
   End
   Begin VB.CommandButton cmdDeleteKey 
      Caption         =   "&Delete Key"
      Height          =   396
      Left            =   165
      TabIndex        =   8
      Top             =   4335
      Width           =   1896
   End
   Begin VB.CommandButton cmdSetValue 
      Caption         =   "&Set Value"
      Height          =   396
      Left            =   2094
      TabIndex        =   7
      Top             =   2898
      Width           =   1896
   End
   Begin VB.CommandButton cmdOpenKey 
      Caption         =   "&Open Key"
      Height          =   396
      Left            =   165
      TabIndex        =   6
      Top             =   3858
      Width           =   1896
   End
   Begin VB.CommandButton cmdDeleteValue 
      Caption         =   "D&elete Value"
      Height          =   396
      Left            =   2094
      TabIndex        =   5
      Top             =   4305
      Width           =   1896
   End
   Begin VB.CommandButton CmdEnumerateKeyNames 
      Caption         =   "&Enumerate Key Names"
      Height          =   396
      Left            =   165
      TabIndex        =   4
      Top             =   2906
      Width           =   1896
   End
   Begin VB.CommandButton cmdEnumerateValueNames 
      Caption         =   "E&numerate Value Names"
      Height          =   396
      Left            =   2094
      TabIndex        =   3
      Top             =   3366
      Width           =   1896
   End
   Begin VB.CommandButton CmdQueryValue 
      Caption         =   "&Query Value"
      Height          =   396
      Left            =   2115
      TabIndex        =   2
      Top             =   3834
      Width           =   1896
   End
   Begin VB.ListBox lst1 
      Height          =   2205
      ItemData        =   "Test_Form1.frx":0000
      Left            =   150
      List            =   "Test_Form1.frx":0002
      TabIndex        =   1
      Top             =   165
      Width           =   3825
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clea&r"
      Height          =   396
      Left            =   165
      TabIndex        =   0
      Top             =   2430
      Width           =   3825
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************'
'------------------------------------------------------'
' Project: EzRegAPI v1.0.23
'
' Date: July-28-2001
'
' Programmer: Antonio Ramirez Cobos
'
' Module: TestForm
'
' Description: Test application for EzRegAPI control. The control
'              is tested for most of the control's functionality but
'              the Association/removal of file types. This only works
'              on Windows 9X platforms, but it could easily be changed
'              to work on NT or 2000 [like to be challenged?]
'
'              Please be careful with the examples of the included help
'              file, as the help file contains errata on some of its
'              examples [better stick with the examples shown on this
'              application]. The control was upgraded not long ago and
'              I didn't had the time to update its help file too.
'
'              From the Author:
'              'cause I consider myself in a continuous learning
'              path with no end on programming, please, if you
'              can improve this program [I am sure you will]
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
Private Sub cmdClear_Click()
 lst1.Clear
End Sub

Private Sub CmdCreateKey_Click()
'The following will create a key under the HKEY_CURRENT_USER root key with the name of the application title:
On Error Resume Next
Dim hKey As Long
lst1.Clear
' Create the key and get a handle to it
hKey = EzRegApi1.CreateKey(EzRegApi1.HkeyCurrentUser, App.Title, EzRegApi1.TotalAccess)
If Err.Number <> 0 Then
    lst1.AddItem "Error..."
    Exit Sub
End If
lst1.AddItem App.Title & " on HKEY_CURRENT_USER created"
' Close the key
EzRegApi1.CloseKey hKey


End Sub

Private Sub cmdDeleteKey_Click()
On Error Resume Next
lst1.Clear
EzRegApi1.DeleteKey EzRegApi1.HkeyCurrentUser, App.Title
If Err.Number <> 0 Then
    lst1.AddItem "Error..."
Else
    lst1.AddItem App.Title & " key successfully deleted"
End If
End Sub

Private Sub cmdDeleteValue_Click()
On Error Resume Next
lst1.Clear
Dim lKey As Long
lKey = EzRegApi1.OpenKey(EzRegApi1.HkeyCurrentUser, App.Title, EzRegApi1.TotalAccess)
EzRegApi1.DeleteValue lKey, "Width"
If Err.Number <> 0 Then
    lst1.AddItem "Error..."
Else
    lst1.AddItem "Width value successfully deleted"
End If
End Sub

Private Sub CmdEnumerateKeyNames_Click()
On Error Resume Next
'The following will print the keys in HKEY_CURRENT_USER:
lst1.Clear
Do
    Keyname = EzRegApi1.EnumKeyNames(EzRegApi1.HkeyCurrentUser)
    If Err.Number <> 0 Then
        lst1.AddItem "Error..."
        Exit Do
    End If
    lst1.AddItem Keyname
Loop While Keyname <> ""

End Sub

Private Sub cmdEnumerateValueNames_Click()
Dim valueName As String, hKey As Long
hKey = EzRegApi1.OpenKey(EzRegApi1.HkeyCurrentUser, App.Title, EzRegApi1.TotalAccess)
lst1.Clear
Do
    valueName = EzRegApi1.EnumValueNames(hKey)
    If Err.Number <> 0 Then
        lst1.AddItem "Error..."
        Exit Do
    End If
    lst1.AddItem valueName
Loop While valueName <> ""
End Sub

Private Sub cmdOpenKey_Click()
Dim hKey As Long
lst1.Clear
On Error Resume Next
' Create the key and get a handle to it
hKey = EzRegApi1.OpenKey(EzRegApi1.HkeyCurrentUser, _
                            App.Title, EzRegApi1.TotalAccess)
If Error.Number <> 0 Then
    lst1.AddItem "Error..."
Else
    lst1.AddItem "Key Opened with handle: " & CStr(hKey)
End If
' Close the key
EzRegApi1.CloseKey hKey
End Sub

Private Sub CmdQueryValue_Click()
'The following code will open a key named
'with the application title and query the
'data stored in the values named 'Width'
On Error Resume Next

Dim lWidth As Long
Dim hKey As Long
lst1.Clear
' Open the Key
hKey = EzRegApi1.OpenKey(EzRegApi1.HkeyCurrentUser, App.Title, EzRegApi1.TotalAccess)
' Get the data stored in 'Width' and 'Height' values
lWidth = CLng(EzRegApi1.QueryValue(hKey, "Width", regValdata))
If Err.Number <> 0 Then
    lst1.AddItem "Error..."
Else
    lst1.AddItem "Width: " & CStr(lWidth)
End If
' Close the key
EzRegApi1.CloseKey hKey
' Set up the form size ?
'Me.Width = lWidth

End Sub

Private Sub cmdSetValue_Click()
'The following code will set a value named "Width" to a DWORD value of 1000:
On Error Resume Next
Dim hKey As Long, valueData
' Open the key under HKEY_CURRENT_USER with the name of the application title
hKey = EzRegApi1.OpenKey(EzRegApi1.HkeyCurrentUser _
       , App.Title, EzRegApi1.TotalAccess)
' Now set the value
valueData = 1000
EzRegApi1.SetValue hKey, "Width", regDword, valueData
If Err.Number Then
    lst1.AddItem "Error..."
Else
    lst1.AddItem "Value Width set to " & valueData
End If
' Close the key
EzRegApi1.CloseKey hKey

End Sub
