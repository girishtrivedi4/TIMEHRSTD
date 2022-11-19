Attribute VB_Name = "mdlLvUpdate"
'' Leave Updation Module
'' ---------------------

'' This is Module is Dedicated to the Updation of Yearly Leaves.
'' This Module Contains Functions for both Yearly Leave Updation and for Updation of
'' Leaves when a New Employee is Added.

Option Explicit

Public Sub UpdateYearlyLeave(ByVal intYear As Integer, ByRef MSF1 As MSFlexGrid, frm As Form)
On Error GoTo ERR_P         '' Updates the Yearly Leave Files
Dim sngTmp As Single
'' Check for Current Year Files
If Not CheckFiles(intYear) Then
    MsgBox NewCaptionTxt("M6001", adrsMod), vbInformation
    Exit Sub
End If
'' Check for Previous Files
If Not CheckFiles(intYear - 1) Then
    'MsgBox NewCaptionTxt("M6002", adrsMod) & CStr(intYear - 1) & _
    'NewCaptionTxt("00055", adrsMod) & vbCrLf & NewCaptionTxt("M6003", adrsMod), vbInformation, App.EXEName
    'Exit Sub
End If
'' Updating Form Status
MSF1.Visible = True
frm.MSF2.Visible = True
frm.Refresh
'' Check for Empty Leave Master File
Call OpenLeaveMaster
If adrsLeave.RecordCount <= 0 Then GoTo Sub_End

Dim strTmp As String
strTmp = "01-" & MonthName(pVStar.Yearstart, True) & "-" & pVStar.YearSel

'' Start Leave Updation
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select EmpCode,Cat,JoinDate,LeavDate from Empmst order by EmpCode", _
ConMain, adOpenStatic
'' Start Employee Loop
If Not adrsDept1.EOF Then
    MSF1.Redraw = True
    MSF1.Refresh
    frm.Refresh
    
    If SubLeaveFlag = 1 Then    ' 07-11
        Dim ELLeave As String, ELSubLeave As String
        If FieldExists("LvBaL" & Right(CStr(intYear), 2), "EL") Then ELLeave = "EL"
        If ELLeave = "EL" Then
            If FieldExists("LvBaL" & Right(Year(Date), 2), "EN") Then ELSubLeave = ",EN"
            If FieldExists("LvBaL" & Right(Year(Date), 2), "NE") Then ELSubLeave = ELSubLeave & ",NE"
            ELSubLeave = Right(ELSubLeave, Len(ELSubLeave) - 1)
        End If
    End If
    
    Do While Not adrsDept1.EOF
        
        If adrsDept1("leavdate") <= CDate(strTmp) And Not IsNull(adrsDept1("leavdate")) Then GoTo Emp_Loop
        MSF1.TextMatrix(1, 0) = adrsDept1("EmpCode")
        '' Start Leave Loop
        adrsLeave.MoveFirst
        Do While Not adrsLeave.EOF
            If adrsDept1("Cat") <> adrsLeave("Cat") Then GoTo Leave_Loop
            Call FillLeaveDetails(adrsLeave("LvCode"))
            If SubLeaveFlag = 1 And (adrsLeave("LvCode") = "EL") Then GoTo Leave_Loop     ' 07-11
            MSF1.TextMatrix(1, 1) = adrsLeave("LvCode")
            MSF1.TextMatrix(1, 2) = adrsLeave("Leave")
            MSF1.Refresh
            frm.Refresh
            '' Credit Current Years Leaves
            If Not adrsLeave.EOF Then
                If typLvD.blnLvType Then
                    If CreditL(intYear) = 0 Then
                        Exit Sub
                    End If
                End If
                
                MSF1.Refresh
                frm.Refresh
                If typLvD.blnCarry = True Then
                    '' Add Opening Balances of Last Years Leaves
                    '' Last years Balance
                    If Not CheckFiles(intYear - 1) Then
'                        MsgBox NewCaptionTxt("M6002", adrsMod) & CStr(intYear - 1) & _
'                        NewCaptionTxt("00055", adrsMod) & vbCrLf & NewCaptionTxt("M6003", adrsMod), vbInformation, App.EXEName
                    Else
                        sngTmp = ReturnOtherValues(intYear - 1, adrsDept1("EmpCode"))
                    End If
                    sngTmp = GetLeaveQty(sngTmp, _
                    ReturnOtherValues(intYear, adrsDept1("EmpCode")))
                     'end by
                    
                    If sngTmp <> 0 Then
                        sngTmp = RoundedLeave(sngTmp)
                        '' Put Necessary Parameters for LeaveInfo
                        typLvI.sngQty = sngTmp
                        typLvI.strFrom = DateCompStr(GetDateOfDay(1, pVStar.Yearstart, intYear))
                        typLvI.strEntry = typLvI.strFrom
                        typLvI.strTo = typLvI.strFrom
                        '' Add Records to Leave Info
                        If AddLeaveInfo(adrsDept1("EmpCode"), intYear, 1) Then
                            '' Update Leave Balance
                            Call UpdateBalance(intYear, adrsDept1("EmpCode"))
                        End If
                    End If
                End If
            End If
Leave_Loop:
                adrsLeave.MoveNext
            Loop
            If SubLeaveFlag = 1 And ELLeave = "EL" Then     ' 07-11
                Dim strqry As String
                strqry = "select " & ELSubLeave & ",lvbal" & Right(CStr(intYear), 2) & ".empcode from lvbal" & Right(CStr(intYear), 2) & " where empcode='" & adrsDept1("EmpCode") & "'"
                Call UpDateSubLeave("lvbal" & Right(CStr(intYear), 2), ELSubLeave, strqry, ELLeave)
            End If
Emp_Loop:
            adrsDept1.MoveNext
    Loop
End If
'' Flag if Leaves are Updated
ConMain.Execute "Update install set lvupdtyear=" & intYear
MSF1.Visible = False
frm.MSF2.Visible = False
frm.Refresh
Sub_End:
'    MsgBox NewCaptionTxt("M6004", adrsMod), vbInformation
Exit Sub
ERR_P:
    ShowError ("UpdateYearLyLeave :: Leave Updation Module")
    MSF1.Visible = False
End Sub

Private Function CheckFiles(ByVal intYear As Integer) As Boolean
On Error GoTo ERR_P
CheckFiles = True               '' Function To Check for Existence of Yearly Leave File.
If Not FindTable("LvTrn" & Right(CStr(intYear), 2)) Then CheckFiles = False
If Not FindTable("LvInfo" & Right(CStr(intYear), 2)) Then CheckFiles = False
If Not FindTable("LvBal" & Right(CStr(intYear), 2)) Then CheckFiles = False
Exit Function
ERR_P:
    ShowError ("Check FIles :: Leave Update")
    CheckFiles = False
End Function

Private Function GetLeaveQty(ByVal sngQtyL As Single, ByVal sngQtyC As Single) As Single
'' Function to return the Valid Leave Amount
If sngQtyL >= typLvD.sngAccQty - sngQtyC Then
    GetLeaveQty = typLvD.sngAccQty - sngQtyC
Else
    GetLeaveQty = sngQtyL
End If
End Function

Public Sub OpenLeaveMaster()
On Error GoTo ERR_P
'' Function to Open the Leave Master Table
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "Select * from LeavDesc Where LvCode not in('" & pVStar.WosCode & "','" & _
pVStar.HlsCode & "','" & pVStar.PrsCode & "','" & pVStar.AbsCode & "')" _
, ConMain, adOpenStatic
Exit Sub
ERR_P:
    ShowError ("OpenLeaveMaster :: Leave Updation Module")
    Resume Next
End Sub

Public Sub FillLeaveDetails(ByVal strLvCode As String)
On Error GoTo ERR_P
'' Function to Fill LeaveDetails
If Not adrsLeave.EOF Then
    typLvD.blnCarry = IIf(adrsLeave("Lv_Cof") = "Y", True, False)
    typLvD.blnCrImd = IIf(adrsLeave("CreditNow") = "Y", True, False)
    typLvD.blnFullPro = IIf(adrsLeave("FulCredit") = "Y", True, False)
    typLvD.blnLvType = IIf(adrsLeave("Type") = "Y", True, False)
    typLvD.sngAccQty = IIf(IsNull(adrsLeave("Lv_Acumul")), 0, adrsLeave("Lv_Acumul"))
    typLvD.sngQty = IIf(IsNull(adrsLeave("Lv_Qty")), 0, adrsLeave("Lv_Qty"))
    typLvD.strCat = adrsLeave("Cat")
    typLvD.strLvCode = adrsLeave("LvCode")
    typLvD.strLvName = adrsLeave("Leave")
End If
Exit Sub
ERR_P:
    ShowError ("FillLeaveDetails :: Leave Updation Module")
End Sub

Public Function AddLeaveInfo(ByVal strEmpCode As String, ByVal intYear As Integer, _
ByVal bytTRCD As Byte) As Boolean       '' Function to Add the Leaves to Leave Info File
On Error GoTo ERR_P
AddLeaveInfo = True

    ConMain.Execute "insert into LvInfo" & Right(CStr(intYear), 2) & "(" & _
    "EmpCode,Trcd,FromDate,ToDate,LCode,Days,EntryDate) Values('" & strEmpCode & _
    "'," & bytTRCD & "," & strDTEnc & typLvI.strFrom & strDTEnc & "," & strDTEnc & _
    typLvI.strTo & strDTEnc & ",'" & typLvD.strLvCode & "'," & typLvI.sngQty & "," & _
    strDTEnc & typLvI.strEntry & strDTEnc & ")"

Exit Function
ERR_P:
    ShowError ("AddLeaveInfo :: Leave Updation Module")
    AddLeaveInfo = False
End Function

Private Function CreditL(ByVal intYear As Integer) As Byte
On Error GoTo ERR_P         '' Function to Deal With Carry ForWard Leaves
Dim sngTmp As Single, sngTmp1 As Single
CreditL = 1
If Not FieldExists("LvBaL" & Right(CStr(intYear), 2), typLvD.strLvCode) Then
    MsgBox NewCaptionTxt("M6005", adrsMod), vbCritical
    CreditL = 0        '' For Critical Error
    Exit Function
End If
If Not CheckFiles(intYear - 1) Then
Else
If Not FieldExists("LvTrn" & Right(CStr(intYear - 1), 2), typLvD.strLvCode) Then
    ''  because it was not updating new leave code defined the next year
    'CreditL = 0        '' For Critical Error
    'Exit Function
End If
End If
If typLvD.blnFullPro = True Then           '' If Full Leave
    sngTmp = typLvD.sngQty
    If sngTmp > 0 Then
        sngTmp = RoundedLeave(sngTmp)
        '' Put Necessary Parameters for LeaveInfo
        typLvI.sngQty = sngTmp
        typLvI.strFrom = DateCompStr(GetDateOfDay(1, pVStar.Yearstart, intYear))
        typLvI.strEntry = typLvI.strFrom
        typLvI.strTo = typLvI.strFrom
        '' Add Records to Leave Info
        If AddLeaveInfo(adrsDept1("EmpCode"), intYear, 2) Then
            '' Update Leave Balance
            Call UpdateBalance(intYear, adrsDept1("EmpCode"))
        End If
    End If
Else    '' If Proportionate
    If pVStar.Yearstart = 1 Then
        sngTmp1 = 12
    Else
        sngTmp1 = pVStar.Yearstart - 1      '' Get the End Month Name
    End If
    '' Get the Work days
    sngTmp1 = DateDiff("d", DateCompDate(adrsDept1("JoinDate")), _
              GetDateOfDay(GetENums(sngTmp1, intYear), sngTmp1, intYear - 1)) 'A
    If sngTmp1 < 0 Then sngTmp = 0
    If sngTmp1 > 365 Then sngTmp1 = 365
    sngTmp = ReturnOtherValues(intYear - 1, adrsDept1("EmpCode"), 2) 'B,Last Years Paid Days
    '' Leaves to be Credited
    
        sngTmp = typLvD.sngQty * sngTmp / sngTmp1

    If sngTmp > 0 Then
        sngTmp = RoundedLeave(sngTmp)
        '' Put Necessary Parameters for LeaveInfo
        typLvI.sngQty = sngTmp
        typLvI.strFrom = DateCompStr(GetDateOfDay(1, pVStar.Yearstart, intYear))
        typLvI.strEntry = typLvI.strFrom
        typLvI.strTo = typLvI.strFrom
        '' Add Records to Leave Info
        If AddLeaveInfo(adrsDept1("EmpCode"), intYear, 2) Then
            Call UpdateBalance(intYear, adrsDept1("EmpCode"))
        End If
    End If
End If
Exit Function
ERR_P:
    ShowError ("CreditL :: Leave Updation Module ")
    CreditL = 0
    Resume Next
End Function

Private Function ReturnOtherValues(ByVal intYear As Integer, strEmpCode As String, _
Optional bytLast As Byte = 1) As Single     '' Retutns other Values
On Error GoTo ERR_P
Select Case bytLast
    Case 1      '' Specified Years Leave Balance
        If adrsPaid.State = 1 Then adrsPaid.Close
        If FieldExists("LvBal" & Right(CStr(intYear), 2), typLvD.strLvCode) Then
            If adrsPaid.State = 1 Then adrsPaid.Close
            adrsPaid.Open "Select " & typLvD.strLvCode & " from LvBal" & _
            Right(CStr(intYear), 2) & " Where EmpCode='" & strEmpCode & _
            "'", ConMain
            If Not adrsPaid.EOF Then
                If IsNull(adrsPaid(0)) Or IsEmpty(adrsPaid(0)) Then
                    ReturnOtherValues = 0
                Else
                    ReturnOtherValues = adrsPaid(0)
                End If
            Else
                ReturnOtherValues = 0
            End If
        Else
            ReturnOtherValues = 0
        End If
    Case 2      '' Last Years Paid Days
    If Not CheckFiles(intYear - 1) Then
'                        MsgBox NewCaptionTxt("M6002", adrsMod) & CStr(intYear - 1) & _
'                        NewCaptionTxt("00055", adrsMod) & vbCrLf & NewCaptionTxt("M6003", adrsMod), vbInformation, App.EXEName
    Else
        If adrsPaid.State = 1 Then adrsPaid.Close
        adrsPaid.Open "Select Sum(Paiddays) from LvTrn" & Right(CStr(intYear), 2) & _
        " Where EmpCode='" & strEmpCode & "'", ConMain
        If Not (adrsPaid.EOF And adrsPaid.BOF) Then
            If IsNull(adrsPaid(0)) Or IsEmpty(adrsPaid(0)) Then
                ReturnOtherValues = 0
            Else
                ReturnOtherValues = adrsPaid(0)
            End If
        Else
            ReturnOtherValues = 0
        End If
    End If
    Case 3
End Select
Exit Function
ERR_P:
    ShowError ("ReturnOtherValues :: Leave Updation Module")
    ReturnOtherValues = 0
End Function

Private Function GetDateOfDay(ByVal bytDay As Byte, ByVal bytMonth As String, _
ByVal intYear As Integer) As String        '' Function to make Date
On Error GoTo ERR_P
Select Case bytDateF
    Case 1      '' American (MM/DD/YY)
        GetDateOfDay = Format(bytMonth, "00") & "/" & Format(bytDay, "00") & _
        "/" & intYear
    Case 2      '' British  (DD/MM/YY)
        GetDateOfDay = Format(bytDay, "00") & "/" & Format(bytMonth, "00") & _
        "/" & intYear
End Select
Exit Function
ERR_P:
    ShowError ("Get Date of the Day Leave Update::")
End Function

Private Function GetENums(ByVal bytTmp As String, _
ByVal intYear As String) As Byte        '' Function to get the End Day of the Specified Month & Year
Select Case bytTmp
    Case 1
        GetENums = 31
    Case 2
        If LeapOrNotUpd(intYear) Then     '' Check out if Leap Year
            GetENums = 29
        Else
            GetENums = 28
        End If
    Case 3
        GetENums = 31
    Case 4
        GetENums = 30
    Case 5
        GetENums = 31
    Case 6
        GetENums = 30
    Case 7
        GetENums = 31
    Case 8
        GetENums = 31
    Case 9
        GetENums = 30
    Case 10
        GetENums = 31
    Case 11
        GetENums = 30
    Case 12
        GetENums = 31
End Select
End Function

Private Function LeapOrNotUpd(ByVal intYear As Integer) As Boolean   '' Checks if a
If intYear Mod 4 = 0 Then   '' If Divisible by 4
    ' Is it a Century?
    If intYear Mod 100 = 0 Then     '' if Divisible by 100
        ' If a Century, must be Evenly Divisible by 400.
        If intYear Mod 400 = 0 Then     '' If Divisible by 400
            LeapOrNotUpd = True                 '' Leap Year
        Else
            LeapOrNotUpd = False                '' Non-Leap Year
        End If
    Else
        LeapOrNotUpd = True                     '' Leap Year
    End If
Else
    LeapOrNotUpd = False                        '' Non-Leap Year
End If
End Function

Public Sub UpdateNewEmpLeave(ByVal strEmpCode As String, ByVal dttmp As Date, _
ByVal strCatTmp As String, ByVal intYear As Integer)
On Error GoTo ERR_P         '' Procedure to Update New Employees Leaves
'' dtTmp is Employees Join date
Dim sngTmp As Single
'' Start Leave Loop
ConMain.Execute "insert into LvBal" & Right(CStr(intYear), 2) & "(EmpCode) " & _
"Values('" & strEmpCode & "')" 'Insert Employees Name in the Leave Balance File

Do While Not adrsLeave.EOF
    If strCatTmp <> adrsLeave("Cat") Then GoTo Leave_Loop
    Call FillLeaveDetails(adrsLeave("LvCode"))
    sngTmp = 0      '' Reset
    If Not adrsLeave.EOF Then
        If FieldExists("LvBal" & Right(CStr(intYear), 2), typLvD.strLvCode) Then
            If typLvD.blnCrImd = True Then
                Select Case DateDiff("m", Year_Start, dttmp)
                    Case Is <= 0        '' Joined Before Year Start Date
                        sngTmp = 12
                    Case 1 To 11        '' Joined in Current Year
                        sngTmp = 12 + pVStar.Yearstart - Month(dttmp)
                        If sngTmp > 12 Then sngTmp = sngTmp - 12
                    Case Else           '' Joined in a Futuristic Date
                        sngTmp = 0
                End Select
                ' C*((B-A + 1)/12.)
                ''sngTmp = typLvD.sngQty * ((sngTmp - Month(dtTmp) + 1) / 12)
'                If sngTmp > 0 Then
'                    sngTmp = typLvD.sngQty * (sngTmp / 12)
'                End If
                If sngTmp > 0 Then

                        If adrsLeave!CRMONTHLY = "Y" Then   '  FOR LEAVE UPDATION FOR NEW EMPLOYEE 15-06
                            If Day(dttmp) >= 15 Then
                                sngTmp = typLvD.sngQty / 2
                            Else
                                sngTmp = typLvD.sngQty
                            End If
                        Else
                            sngTmp = typLvD.sngQty * (sngTmp / 12)
                        End If

                End If
                
                sngTmp = GetLeaveQty(sngTmp, ReturnOtherValues(intYear, strEmpCode))
                If sngTmp > 0 Then
                    sngTmp = RoundedLeave(sngTmp)
                    '' Put Necessary Parameters for LeaveInfo
                    typLvI.sngQty = sngTmp
                    
                    If dttmp < Year_Start Then
                        typLvI.strFrom = DateSaveIns(Year_Start)
                        typLvI.strTo = typLvI.strFrom
                    Else
                        typLvI.strFrom = dttmp
                        typLvI.strTo = typLvI.strFrom
                    End If
                    ''
                    typLvI.strEntry = DateSaveIns(Date)
                    '' Add Records to Leave Info
                    If AddLeaveInfo(strEmpCode, intYear, 2) Then
                        Call UpdateBalance(intYear, strEmpCode)
                    End If
                End If
            End If
        Else
            '''Msgbox "Leave " & typLvD.strLvCode & " not Found in the Leave Balance File" & _
            " Leave Could not be Updated", vbExclamation, App.EXEName
        End If
    End If
Leave_Loop:
    adrsLeave.MoveNext
Loop
Exit Sub
ERR_P:
    ShowError ("UpdateNewEmpLeave :: Leave Updation Module")
    'Resume Next
End Sub

Private Sub UpdateBalance(ByVal intYear As Integer, strEmpCode As String)
On Error GoTo ERR_P
'' Where Leave is not Null
ConMain.Execute "Update LvBal" & Right(CStr(intYear), 2) & _
" Set " & typLvD.strLvCode & "=" & typLvD.strLvCode & "+" & typLvI.sngQty _
& " Where EmpCode='" & strEmpCode & "' and " & typLvD.strLvCode & " is Not Null"
'' Where Leave is Null
ConMain.Execute "Update LvBal" & Right(CStr(intYear), 2) & _
" Set  " & typLvD.strLvCode & "=" & typLvI.sngQty & " Where EmpCode='" & _
strEmpCode & "' and " & typLvD.strLvCode & " is Null"
Exit Sub
ERR_P:
    ShowError ("Update Balance :: Leave Update")
End Sub

Public Function RoundedLeave(ByVal sngTmp As Single) As Single
sngTmp = Round(sngTmp, 2)
Select Case sngTmp - Fix(sngTmp)
    Case 0 To 0.25
        RoundedLeave = Fix(sngTmp)
    Case 0.26 To 0.74
        RoundedLeave = Fix(sngTmp) + 0.5
    Case 0.75 To 0.99
        RoundedLeave = Fix(sngTmp) + 1
End Select
End Function

'Added by  07-11
Public Function UpDateSubLeave(ByVal LvBalTbl As String, ByVal SubLv As String, ByVal Qry As String, ByVal MainLv As String)
Dim arrSubLv() As String
Dim Sum As Single
Dim j As Integer

arrSubLv = Split(SubLv, ",")
If adrsASC.State = 1 Then adrsASC.Close
adrsASC.Open Qry, ConMain, adOpenStatic
Do While Not (adrsASC.EOF)
    For j = 0 To UBound(arrSubLv)
        Sum = Sum + IIf(IsNull(adrsASC.Fields(arrSubLv(j))), 0, adrsASC.Fields(arrSubLv(j)))
    Next
    ConMain.Execute "Update " & LvBalTbl & " Set " & _
    MainLv & "=" & Sum & " where " & LvBalTbl & ".Empcode='" & adrsASC.Fields("Empcode") & "'"
    Sum = 0
    adrsASC.MoveNext
Loop
End Function
