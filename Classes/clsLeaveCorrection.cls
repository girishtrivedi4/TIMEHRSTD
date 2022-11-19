VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeaveCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsLeaveCorrection
' DateTime  : 14/Aug/2008 14:05
' Author    :
' Purpose   : This class used to correct leave after updated.
'---------------------------------------------------------------------------------------
'MIS2007DF04
Option Explicit
'This variable used to stored leave details.
Private LeaveInfo As clsLeaveDetails
'This variable used to stored leave balance table name.
Private LeaveBalName As String
'This variable used to stored leave Info table name.
Private LeaveInfoName As String
'This variable used to stored value of Opening date of leave.
Private FromDate As String
'---------------------------------------------------------------------------------------
' Procedure : GetLeaveDetails
' DateTime  : 14/Aug/2008 09:59
' Author    :
' Purpose   : This function used to get leave details from backend and
'                stored in LvDesc user defined type
' Parameter : 1)mLeaveCode:which leave code detail u want.
' Return    : LvDesc
'---------------------------------------------------------------------------------------
'
Public Function GetLeaveDetails(mLeaveCode As String) _
    As clsLeaveDetails
    Dim adrTemp As Recordset
    Dim LeaveDetails As New clsLeaveDetails
On Error GoTo GetLeaveDeatails_Error
    Set adrsTemp = OpenRecordSet("SELECT * FROM Leavdesc WHERE lvcode='" & _
        mLeaveCode & "'")
    LeaveDetails.blnCarry = IIf(adrsTemp("Lv_Cof") = "Y", True, False)
    LeaveDetails.blnCrImd = IIf(adrsTemp("CreditNow") = "Y", True, False)
    LeaveDetails.blnFullPro = IIf(adrsTemp("FulCredit") = "Y", True, False)
    LeaveDetails.blnLvType = IIf(adrsTemp("Type") = "Y", True, False)
    LeaveDetails.sngAccQty = IIf(IsNull(adrsTemp("Lv_Acumul")), 0, adrsTemp("Lv_Acumul"))
    LeaveDetails.sngQty = IIf(IsNull(adrsTemp("Lv_Qty")), 0, adrsTemp("Lv_Qty"))
    LeaveDetails.strCat = adrsTemp("Cat")
    LeaveDetails.strLvCode = adrsTemp("LvCode")
    LeaveDetails.strLvName = adrsTemp("Leave")
    Set GetLeaveDetails = LeaveDetails
    Set adrsTemp = Nothing
    Set LeaveDetails = Nothing
On Error GoTo 0
Exit Function
GetLeaveDeatails_Error:
   If Erl = 0 Then
      ShowError "Error in procedure GetLeaveDeatails of Class Module clsLeaveCorrection"
   Else
      ShowError "Error in procedure GetLeaveDeatails of Class Module clsLeaveCorrection And Line:" & Erl
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLeaveBalOfYear
' DateTime  : 14/Aug/2008 11:06
' Author    :
' Purpose   : This function Used to Get leave balance of given year and given leave code.
' Return    : Single
'---------------------------------------------------------------------------------------
'
Public Function GetLeaveBalOfYear(mLeaveCode As String, _
    mYear As Integer, mEmpcode As String) As Single
    Dim strTableName As String
    Dim sngBal As Single
On Error GoTo GetLeaveBalOfYear_Error
    strTableName = LeaveFile(mYear, LvBal)
    If Not FindTable(strTableName) Then
        MsgBox strTableName & " This file not present", vbInformation, _
            "GetLeaveBalOfYear in clsLeaveCorrection"
        Exit Function
    End If
    If Not FieldExists(strTableName, mLeaveCode) Then
        MsgBox mLeaveCode & " This field not exit in " & _
            strTableName, vbInformation, _
            "GetLeaveBalOfYear in clsLeaveCorrection"
        Exit Function
    End If
    sngBal = ExecScalar("SELECT " & mLeaveCode & " FROM " & _
        strTableName & " WHERE empcode='" & mEmpcode & "'")
    GetLeaveBalOfYear = sngBal
On Error GoTo 0
Exit Function
GetLeaveBalOfYear_Error:
   If Erl = 0 Then
      ShowError "Error in procedure GetLeaveBalOfYear of Class Module clsLeaveCorrection"
   Else
      ShowError "Error in procedure GetLeaveBalOfYear of Class Module clsLeaveCorrection And Line:" & Erl
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : LeaveFile
' DateTime  : 14/Aug/2008 11:14
' Author    :
' Purpose   : This function used to Get Leave file Name.
' Return    : String
'---------------------------------------------------------------------------------------
'
Public Function LeaveFile(mYear As Integer, LvFile As eLeaveFile) As String
On Error GoTo LeaveFile_Error
    Select Case LvFile
        Case LvBal
            LeaveFile = "LVBAL" & Right(mYear, 2)
        Case LvInfo
            LeaveFile = "LVINFO" & Right(mYear, 2)
        Case LvTrn
            LeaveFile = "LVTRN" & Right(mYear, 2)
    End Select
On Error GoTo 0
Exit Function
LeaveFile_Error:
   If Erl = 0 Then
      ShowError "Error in procedure LeaveFile of Class Module clsLeaveCorrection"
   Else
      ShowError "Error in procedure LeaveFile of Class Module clsLeaveCorrection And Line:" & Erl
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : LeaveCorrection
' DateTime  : 14/Aug/2008 11:30
' Author    :
' Purpose   : This procedure is start up of this class
' Return    :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
On Error GoTo Class_Initialize_Error
    Set LeaveInfo = GetLeaveDetails("SL")
    LeaveBalName = LeaveFile(Val(pVStar.YearSel), LvBal)
    LeaveInfoName = LeaveFile(Val(pVStar.YearSel), LvInfo)
    FromDate = DateCompStr(GetDateOfDay(1, MonthName(pVStar.Yearstart), pVStar.YearSel))
On Error GoTo 0
Exit Sub
Class_Initialize_Error:
   If Erl = 0 Then
      ShowError "Error in procedure Class_Initialize of Class Module clsLeaveCorrection"
   Else
      ShowError "Error in procedure Class_Initialize of Class Module clsLeaveCorrection And Line:" & Erl
   End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateBalance
' DateTime  : 14/Aug/2008 12:19
' Author    :
' Purpose   : This is main function.
' Return    :
'---------------------------------------------------------------------------------------
'
'Special Comment: for this Used standard statistics formula:
'                   Current Value - Wrong Value + Currect Value.
Public Sub UpdateBalance(Empcode As String)
    'This variable used to stored Empcode.
    Dim mEmpcode As String
    Dim sngWrong As Single    'Wrong Value
    Dim sngCurrent As Single    'Current Value
    Dim sngCorrect As Single    'Correct Value
On Error GoTo UpdateBalance_Error
    mEmpcode = Empcode
    '(Current Value - Wrong Value) This function Execute 1st step
    sngWrong = EntryDelInLvInfo(mEmpcode, FromDate, Opening) 'GetLeaveBalOfYear(LeaveInfo.strLvCode, pVStar.YearSel - 1, mEmpcode)
    'To get balance of previous year
    sngCurrent = GetLeaveBalOfYear(LeaveInfo.strLvCode, _
       Val(pVStar.YearSel) - 1, mEmpcode)
    sngCorrect = GetLeaveQty(sngCurrent, LeaveInfo.sngQty)
    '(+ Currect Value.) this line execute 3rd value
    ExecNonQuery "UPDATE " & LeaveBalName & " SET " & LeaveInfo.strLvCode & _
        "=" & LeaveInfo.strLvCode & "+(" & sngCorrect & ") WHERE empcode='" & _
        mEmpcode & "'", "UpdateBalance"
    ExecNonQuery "INSERT INTO " & LeaveInfoName & " (empcode,trcd,fromdate,todate," & _
    "lcode,days,entrydate) VALUES('" & mEmpcode & "'," & Opening & " ," & _
    strDTEnc & "" & FromDate & "" & strDTEnc & "," & _
    strDTEnc & "" & FromDate & "" & strDTEnc & ",'" & _
    LeaveInfo.strLvCode & "'," & sngCorrect & "," & _
    strDTEnc & "" & FromDate & "" & strDTEnc & ")"
On Error GoTo 0
Exit Sub
UpdateBalance_Error:
   If Erl = 0 Then
      ShowError "Error in procedure UpdateBalance of Class Module clsLeaveCorrection"
   Else
      ShowError "Error in procedure UpdateBalance of Class Module clsLeaveCorrection And Line:" & Erl
   End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EntryDelInLvInfo
' DateTime  : 14/Aug/2008 14:00
' Author    :
' Purpose   : This function used to minus wrong value from current value.
' Return    : Single
'---------------------------------------------------------------------------------------
'
Private Function EntryDelInLvInfo(mEmpcode As String, _
    mFromDate As String, trcd As eTrcd) As Single
    'To Get wrong value.
On Error GoTo EntryDelInLvInfo_Error
    EntryDelInLvInfo = ExecScalar("SELECT days FROM " & LeaveInfoName & _
    " WHERE empcode='" & mEmpcode & "' AND fromdate=" & strDTEnc & _
    "" & DateCompStr(mFromDate) & "" & strDTEnc & " AND trcd=" & _
    trcd & _
    " AND lcode='" & LeaveInfo.strLvCode & "'", NumericD, "EntryDelInLvInfo")
    'to deleted wrong entry.
    ExecNonQuery "DELETE FROM " & LeaveInfoName & _
    " WHERE empcode='" & mEmpcode & "' AND fromdate=" & strDTEnc & _
    "" & DateCompStr(mFromDate) & "" & strDTEnc & " AND trcd=" & _
    trcd & " AND lcode='" & LeaveInfo.strLvCode & "'", "EntryDelInLvInfo"
    'to update lvbal
    ExecNonQuery "UPDATE " & LeaveBalName & " SET " & LeaveInfo.strLvCode & _
        "=" & LeaveInfo.strLvCode & "-(" & EntryDelInLvInfo & _
        ") WHERE empcode='" & mEmpcode & "'", "EntryDelInLvInfo"
On Error GoTo 0
Exit Function
EntryDelInLvInfo_Error:
   If Erl = 0 Then
      ShowError "Error in procedure EntryDelInLvInfo of Class Module clsLeaveCorrection"
   Else
      ShowError "Error in procedure EntryDelInLvInfo of Class Module clsLeaveCorrection And Line:" & Erl
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLeaveQty
' DateTime  : 14/Aug/2008 14:01
' Author    :
' Purpose   : Function to return the Valid Leave Amount
' Return    : Single
'---------------------------------------------------------------------------------------
'
Private Function GetLeaveQty(ByVal sngQtyL As Single, ByVal sngQtyC As Single) As Single
On Error GoTo GetLeaveQty_Error
    If sngQtyL >= LeaveInfo.sngAccQty - sngQtyC Then
        GetLeaveQty = LeaveInfo.sngAccQty - sngQtyC
    Else
        GetLeaveQty = sngQtyL
    End If
On Error GoTo 0
Exit Function
GetLeaveQty_Error:
   If Erl = 0 Then
      ShowError "Error in procedure GetLeaveQty of Class Module clsLeaveCorrection"
   Else
      ShowError "Error in procedure GetLeaveQty of Class Module clsLeaveCorrection And Line:" & Erl
   End If
End Function
