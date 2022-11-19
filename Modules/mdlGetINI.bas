Attribute VB_Name = "mdlGetINI"
'---------------------------------------------------------------------------------------
' Module    : mdlGetINI
' DateTime  : 20/06/07 14:10
' Author    : jagdish
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetINI(ApplicationName As String, Key As String, Optional mDefault As String, Optional mFilename As String) As String
    Dim mLen As String
    Dim mstr As String * 3000
    If mFilename = "" Then mFilename = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & App.EXEName & ".INI"
    mLen = GetPrivateProfileString(ApplicationName, Key, mDefault, mstr, 3000, mFilename)
    GetINI = Mid(mstr, 1, mLen)
End Function

Public Function WriteINI(ApplicationName As String, Key As String, Value As String, Optional mFilename As String) As String
    Dim mLen As String
    Dim mstr As String * 255
    If mFilename = "" Then mFilename = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & App.EXEName & ".ini"
    WritePrivateProfileString ApplicationName, Key, Value, mFilename
End Function
