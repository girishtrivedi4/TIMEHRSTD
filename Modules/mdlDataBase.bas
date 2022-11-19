Attribute VB_Name = "mdlDataBase"
'---------------------------------------------------------------------------------------
' Module    : mdlDataBase
' DateTime  : 11/06/2008 14:01
' Author    :
' Purpose   : To perform all common function related to database
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : OpenRecordSet
' DateTime  : 11/06/2008 14:10
' Author    :
' Purpose   : Open Record Set
' Pre       : Query string and optional connection object
' Post      : Making RecordSet as per query string
' Return    : ADODB.Recordset
'---------------------------------------------------------------------------------------
'
Public Function OpenRecordSet(strquery As String, Optional _
    Conn As ADODB.Connection) As ADODB.Recordset

On Error GoTo OpenRecordSet_Error
    If Conn Is Nothing Then ''MIS connection string
        Set OpenRecordSet = ConMain.Execute(strquery)
    Else
        ''for other connection
    End If
On Error GoTo 0
Exit Function
OpenRecordSet_Error:
   ShowError "Error in procedure OpenRecordSet of Module mdlDataBase"
'Resume Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExecScalar
' DateTime  : 11/06/2008 14:12
' Author    :
' Purpose   : To get single value from database
' Pre       : Query string and Acpected Datatype of return value Bydefualt is string
' Post      : Getting single value from database
' Return    : Variant
'---------------------------------------------------------------------------------------
'
Public Function ExecScalar(strquery As String, _
        Optional OutPutDT As DataType = StringD, _
        Optional CallFrom As String = "Common")
    Dim adrsTemp As New ADODB.Recordset
On Error GoTo ExecScalar_Error
    Set adrsTemp = OpenRecordSet(strquery)
    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
        Select Case OutPutDT
            Case StringD
                ExecScalar = FilterNull(adrsTemp.Fields(0))
            Case NumericD
                ExecScalar = FilterNull(adrsTemp.Fields(0), NumericD)
        End Select
    End If
On Error GoTo 0
Exit Function
ExecScalar_Error:
   If Erl = 0 Then
      ShowError "Error in procedure ExecScalar of Module mdlDataBase Calling from " & CallFrom
   Else
      ShowError "Error in procedure ExecScalar of Module mdlDataBase And Line:" & Erl & " Calling from " & CallFrom
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExecNonQuery
' DateTime  : 14/Aug/2008 12:04
' Author    :
' Purpose   : This function used to fire query which did not return any value
' Return    : Variant
'---------------------------------------------------------------------------------------
'
Public Sub ExecNonQuery(strquery As String, _
    Optional CallFrom As String = "Common")
On Error GoTo ExecNonQuery_Error
    ConMain.Execute (strquery)
On Error GoTo 0
Exit Sub
ExecNonQuery_Error:
   If Erl = 0 Then
      ShowError "Error in procedure ExecNonQuery of Module mdlDataBase Calling from " & CallFrom
   Else
      ShowError "Error in procedure ExecNonQuery of Module mdlDataBase And Line:" & Erl & " Calling from " & CallFrom
   End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : FilterNull
' DateTime  : 26/07/2008 11:31
' Author    :
' Purpose   : Use this function to Filter Dbnull value.
' Pre       :
' Post      :
' Return    : String
'---------------------------------------------------------------------------------------
'
Public Function FilterNull(strV As Object, _
    Optional dt As DataType = StringD) As String
On Error GoTo FilterNull_Error
If IsNull(strV) = True Then
    If dt = StringD Then
        FilterNull = ""
    ElseIf dt = NumericD Then
        FilterNull = 0
    End If
Else
    FilterNull = strV
End If
On Error GoTo 0
Exit Function
FilterNull_Error:
   If Erl = 0 Then
      ShowError "Error in procedure FilterNull of Module mdlDataBase"
   Else
      ShowError "Error in procedure FilterNull of Module mdlDataBase And Line:" & Erl
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ParseQuery
' DateTime  : 26/07/2008 12:57
' Author    :
' Purpose   : This function use to Parse Query based on backed selection. When u r refer this function please also update file based on keyword u required.
' Pre       :
' Post      :
' Return    : String
'---------------------------------------------------------------------------------------
'
Public Function ParseQuery(mQuery As String) As String
On Error GoTo ParseQuery_Error
    ParseQuery = Replace(mQuery, "~TEXT~", _
        GetKeyWord("~TEXT~", "VARCHAR"))
    ParseQuery = Replace(ParseQuery, "~SMALLINT~", _
        GetKeyWord("~SMALLINT~", "SMALLINT"))
On Error GoTo 0
Exit Function
ParseQuery_Error:
   If Erl = 0 Then
      ShowError "Error in procedure ParseQuery of Module mdlDataBase"
   Else
      ShowError "Error in procedure ParseQuery of Module mdlDataBase And Line:" & Erl
   End If
End Function

Public Function GetKeyWord(mKeyInput As String, Optional mDefualt As String) As String
On Error GoTo GetKeyWord_Error
    GetKeyWord = GetINIString(CStr(bytBackEnd), mKeyInput, App.path & _
                "\DBKeyWord.ini", mDefualt)
On Error GoTo 0
Exit Function
GetKeyWord_Error:
   If Erl = 0 Then
      ShowError "Error in procedure GetKeyWord of Module mdlDataBase"
   Else
      ShowError "Error in procedure GetKeyWord of Module mdlDataBase And Line:" & Erl
   End If
End Function

