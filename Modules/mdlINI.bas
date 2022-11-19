Attribute VB_Name = "mdlINI"
Option Explicit
'Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#If Win16 Then
   Declare Function WritePrivateProfileString Lib "KERNEL" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal FileName As String) As Integer
   Declare Function GetPrivateProfileString Lib "KERNEL" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal ReturnString As String, ByVal NumBytes As Integer, ByVal FileName As String) As Integer
#ElseIf Win32 Then
   Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal FileName As String) As Long
   Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal ReturnString As String, ByVal NumBytes As Long, ByVal FileName As String) As Long
#End If


Public Function WriteINIString(strSection As String, strKeyName As String, strValue As String, strFile As String) As Long
  Dim lngStatus As Long
  lngStatus& = WritePrivateProfileString(strSection, strKeyName, strValue, strFile)
  WriteINIString& = (lngStatus& <> 0)
End Function

Public Function GetINIString(strSection As String, strKeyName As String, strFile As String, Optional strDefault As String = "") As String
  Dim strBuffer As String * 256, lngSize As Long
  lngSize& = GetPrivateProfileString(strSection$, strKeyName$, strDefault$, strBuffer$, 256, strFile$)
  GetINIString$ = Left$(strBuffer$, lngSize&)
End Function
