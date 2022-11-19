Attribute VB_Name = "mdlRotation"


Option Explicit
'' Intrinsic Variables
Private strRotType As String
Private bytTotalShifts As Byte          '' For Total Number of Shifts (Except WD)
Private bytTotalSkips As Byte           '' For Total Number of Skips (Except WD)
Private bytTotalSkipDays As Byte        '' For Total Number of Skips Days (Except WD)
Private bytDaysLeft As Byte             '' For Total Number of Days Left to Skip (Except WD)
Private strPattern() As String          '' For Rotation Pattern
Private strShifts() As String           '' For Shift Pattern
Private strShiftsAll(1 To 31) As String '' For the Shifts to be Inserted into the .Shf File


Private Enum WeekDays
    sun = 2
    mon
    Tue
    wed
    thu
    fri
    sat
End Enum

Private Type StartEndIndex
    StartIndex As Byte
    EndIndex As Byte
End Type
Public Sub WeekDaysShifts1(ByVal strEmpCode As String, ByVal strMonth As String, _
ByVal strYear As String)        '' Procedure to Assign Shift if Rotation type is 'W'
On Error GoTo ERR_P             '' Unused Old Function
Dim dttmp As Date, strTmp As String
Call SetShiftsAll       '' Set the Shifts Array to Blank
typPA.bytWeek = 0
typPA.bytShift = 0
'' Start for the First Day
strShiftsAll(typSENum.bytStart) = strShifts(typPA.bytShift)         '' Put the First Shift
dttmp = DateCompDate(GetDateOfDay(typSENum.bytStart, strMonth, strYear))
If GetWeekDayName(dttmp, 2) = UCase(strPattern(typPA.bytWeek)) Then '' If Next WeekDay then
    typPA.bytWeek = typPA.bytWeek + 1                               '' Move Patterns Pointer
    If typPA.bytWeek > UBound(strPattern) Then typPA.bytWeek = 0    '' Reset if Exceeds
End If
typSENum.bytStart = typSENum.bytStart + 1
Do While typSENum.bytStart <= typSENum.bytEnd     '' Move to Next Day
    dttmp = DateCompDate(GetDateOfDay(typSENum.bytStart, strMonth, strYear))
    If GetWeekDayName(dttmp, 2) = UCase(strPattern(typPA.bytWeek)) Then    '' If Next WeekDay then
        typPA.bytWeek = typPA.bytWeek + 1                           '' Move Patterns Pointer
        If typPA.bytWeek > UBound(strPattern) Then typPA.bytWeek = 0    '' Reset if Exceeds
        typPA.bytShift = typPA.bytShift + 1                         '' Move Shifts Pointer
        If typPA.bytShift > UBound(strShifts) Then typPA.bytShift = 0   '' Reset if Exceeds
    End If
    strShiftsAll(typSENum.bytStart) = strShifts(typPA.bytShift)     '' Put the Next Shift
    strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = pVStar.WosCode  '' For Week Off
    End If
    strTmp = GetHoliday(dttmp)
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = strTmp    '' For Holiday
    End If
    typSENum.bytStart = typSENum.bytStart + 1       '' Increment Start Date
Loop
Exit Sub
ERR_P:
    ShowError ("WeekDaysShifts :: Rotation Module")
End Sub

Public Sub WeekDaysShifts(ByVal strEmpCode As String, ByVal strMonth As String, _
ByVal strYear As String)        '' Procedure to Assign Shift if Rotation type is 'W'
On Error GoTo ERR_P:
Dim dttmp As Date, dtLast As Date, dtStart As Date, bytTmp As Byte, strTmp As String
'' Get the Shift Start Date
dttmp = typEmpRot.dtShift
'' Make the Shift Start Date of the Current Month
dtStart = GetDateOfDay(typSENum.bytStart, strMonth, strYear)
'' Make the Shift Start Date of the Current Month
dtLast = FdtLdt(MonthNumber(strMonth), strYear, "L")
'' Set Initial Values
Call SetShiftsAll       '' Set the Shifts Array to Blank
typPA.bytWeek = 0
typPA.bytShift = 0

    Dim test As Integer
    test = 0
If dtStart <> typEmpRot.dtShift Then
    For test = 0 To UBound(strPattern) Step 1
        If GetWeekDayName(dttmp, 2) <= UCase(strPattern(typPA.bytWeek)) Then
            typPA.bytWeek = test
            Exit For
        End If
        typPA.bytShift = typPA.bytShift + 1
    Next
Else
    If Month(dttmp) = Month(dtStart) And Year(dttmp) = Year(dtStart) Then
        For test = 0 To UBound(strPattern) Step 1
            If GetWeekDayName(dttmp, 2) <= UCase(strPattern(typPA.bytWeek)) Then
                typPA.bytWeek = test
                Exit For
            End If
           typPA.bytShift = typPA.bytShift + 1
        Next
    End If
End If

typPA.bytShift = typPA.bytShift - IIf(typPA.bytShift > 0, 1, 0)

Do While dttmp <= dtLast
    If GetWeekDayName(dttmp, 2) = UCase(strPattern(typPA.bytWeek)) Then         '' If Next WeekDay then
            typPA.bytWeek = typPA.bytWeek + 1                                   '' Move Patterns Pointer
            If typPA.bytWeek > UBound(strPattern) Then typPA.bytWeek = 0        '' Reset if Exceeds
        If dttmp <> typEmpRot.dtShift Then
            typPA.bytShift = typPA.bytShift + 1                                 '' Move Shifts Pointer
            If typPA.bytShift > UBound(strShifts) Then typPA.bytShift = 0       '' Reset if Exceeds
        End If
    End If
    If dttmp >= dtStart Then
        strShiftsAll(typSENum.bytStart) = strShifts(typPA.bytShift)
        strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
        If strTmp <> "" Then
            strShiftsAll(typSENum.bytStart) = pVStar.WosCode  '' For Week Off
        End If
        strTmp = GetHoliday(dttmp)
        If strTmp <> "" Then
            strShiftsAll(typSENum.bytStart) = strTmp    '' For Holiday
        End If
        typSENum.bytStart = typSENum.bytStart + 1
    End If
    dttmp = dttmp + 1
Loop
Exit Sub
ERR_P:
End Sub
'This function made public from private
Public Function GetDateOfDay(ByVal bytDay As Byte, ByVal strMonth As String, _
strYear As String) As String        '' Function to make Date
On Error GoTo ERR_P
Select Case bytDateF
    Case 1      '' American (MM/DD/YY)
        GetDateOfDay = Format(MonthNumber(strMonth), "00") & "/" & Format(bytDay, "00") & _
        "/" & strYear
    Case 2      '' British  (DD/MM/YY)
        GetDateOfDay = Format(bytDay, "00") & "/" & Format(MonthNumber(strMonth), "00") & _
        "/" & strYear
End Select
Exit Function
ERR_P:
    ShowError ("Get Date of day :: Rotation Module")
End Function

Public Sub GetSENums(ByVal strMonth As String, ByVal strYear As String) '' Get the Start
typSENum.bytStart = 1                                                   '' and End Day Numbers
Select Case MonthNumber(strMonth)
    Case 1, 3, 5, 7, 8, 10, 12
        typSENum.bytEnd = 31
    Case 2
        If LeapOrNotRot(CInt(strYear)) Then     '' Check out if Leap Year
            typSENum.bytEnd = 29
        Else
            typSENum.bytEnd = 28
        End If
    Case 4, 6, 9, 11
        typSENum.bytEnd = 30
End Select
End Sub

Private Function LeapOrNotRot(Optional intYear As Integer) As Boolean   '' Checks if a
If IsMissing(intYear) Then                                              '' Specified Year is
    intYear = CInt(pVStar.YearSel)                                      '' Leap Year or not.
End If
If intYear Mod 4 = 0 Then   '' If Divisible by 4
    ' Is it a Century?
    If intYear Mod 100 = 0 Then     '' if Divisible by 100
        ' If a Century, must be Evenly Divisible by 400.
        If intYear Mod 400 = 0 Then     '' If Divisible by 400
            LeapOrNotRot = True                 '' Leap Year
        Else
            LeapOrNotRot = False                '' Non-Leap Year
        End If
    Else
        LeapOrNotRot = True                     '' Leap Year
    End If
Else
    LeapOrNotRot = False                        '' Non-Leap Year
End If
End Function

Private Function GetWeekDayName(ByVal dtDate As Date, Optional bytCut As Byte = 0) _
As String       '' Returns the Name of the WeekDay,also the Specified Left Characters
Dim strTmp As String
Select Case WeekDay(dtDate)
    Case 1
        strTmp = "SUNDAY"
    Case 2
        strTmp = "MONDAY"
    Case 3
        strTmp = "TUESDAY"
    Case 4
        strTmp = "WEDNESDAY"
    Case 5
        strTmp = "THURSDAY"
    Case 6
        strTmp = "FRIDAY"
    Case 7
        strTmp = "SATURDAY"
End Select
Select Case bytCut
    Case 0          '' If no Number Characters are Specified return whole String.
        GetWeekDayName = strTmp
    Case Else       '' If Specified Number of Characters are to be returned.
        GetWeekDayName = Left(strTmp, bytCut)
End Select
End Function

Public Sub FillArrays()      '' Fills the Array of the Pattern and Shifts
On Error GoTo ERR_P
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select Skp,Pattern,Mon_Oth,tot_Shf,tot_Skp,day_Skp from Ro_Shift Where SCode='" & typEmpRot.strShiftCode & _
"'", ConMain
If Not (adrsDept1.EOF And adrsDept1.BOF) Then
    strRotType = UCase(adrsDept1("Mon_Oth"))            '' Type i.e O,W,D
    strCapSND = strRotType                              '' Assign it Back
    strPattern = Split(adrsDept1("Skp"), ",")           '' Skip Patttern
    strShifts = Split(adrsDept1("Pattern"), ".")        '' Shifts Pattern
    ReDim Preserve strPattern(UBound(strPattern) - 1)   '' Clear Last Empty Element
    ReDim Preserve strShifts(UBound(strShifts) - 1)     '' Clear Last Empty Element
    bytTotalShifts = adrsDept1("tot_Shf")               '' Total Number of Shifts
    bytTotalSkips = adrsDept1("tot_Skp")                '' Total Number of Skips
    bytTotalSkipDays = adrsDept1("day_Skp")             '' Total Number of Skip Days
End If
Exit Sub
ERR_P:
    ShowError ("FillArrays :: Rotation Module")
End Sub

Public Sub FillEmployeeDetails(ByVal strEmpCode As String)     '' Fills Specified Employee
On Error GoTo ERR_P
adrsEmp.MoveFirst
adrsEmp.Find "EmpCode='" & strEmpCode & "'"                     '' Details
If Not adrsEmp.EOF Then
       typEmpRot.strCode = strEmpCode               '' Employee Code
       typEmpRot.strCat = adrsEmp("Cat")            '' Category
       typEmpRot.strLocation = adrsEmp("Location")      ' 18-01
       typEmpRot.strOff = IIf(IsNull(adrsEmp("Off")), "", adrsEmp("Off"))          '' First Week Off
       typEmpRot.strOff2 = IIf(IsNull(adrsEmp("Off2")), "", adrsEmp("Off2"))        '' Second Week Off
       typEmpRot.strOff_1_3 = IIf(IsNull(adrsEmp("Wo_1_3")), "", adrsEmp("Wo_1_3"))   '' First and Third Week Off
       typEmpRot.strOff_2_4 = IIf(IsNull(adrsEmp("Wo_2_4")), "", adrsEmp("Wo_2_4"))     '' Second and Fourth Week Off
       typEmpRot.strShifttype = adrsEmp("STyp")     '' Shift Type
       If typEmpRot.strShifttype = "F" Then         '' Shift Code
            typEmpRot.strShiftCode = adrsEmp("F_Shf")
       Else
            typEmpRot.strShiftCode = adrsEmp("SCode")
       End If
       typEmpRot.dtJoin = adrsEmp("JoinDate")       '' Join Date
       typEmpRot.dtShift = adrsEmp("Shf_Date")       '' Shift Date
       '' Leave Date
       typEmpRot.dtLeave = IIf(IsNull(adrsEmp("LeavDate")), Empty, adrsEmp("LeavDate"))
End If
Exit Sub
ERR_P:
    ShowError ("Fill EMployee Details :: Rotation Module")
End Sub

Private Function GetWeekOff(ByVal strWeekDay As String) As String   '' Returns if Week Off
GetWeekOff = ""                                                     '' on Specified Day
If UCase(strWeekDay) = UCase(typEmpRot.strOff) Then GetWeekOff = typEmpRot.strOff
If UCase(strWeekDay) = UCase(typEmpRot.strOff2) Then GetWeekOff = typEmpRot.strOff2
If UCase(strWeekDay) = UCase(typEmpRot.strOff_1_3) Then
    If AlternateWO(1) = True Then GetWeekOff = typEmpRot.strOff_1_3
End If
If UCase(strWeekDay) = UCase(typEmpRot.strOff_2_4) Then
    If AlternateWO(2) = True Then GetWeekOff = typEmpRot.strOff_2_4
End If
End Function

Private Function GetHoliday(ByVal dttmp As Date) As String          '' returns if Holiday
On Error GoTo ERR_P                                                 '' on the Specified Date
GetHoliday = ""
If adrsDept1.State = 1 Then adrsDept1.Close
    adrsDept1.Open "Select Cat from Holiday Where " & strKDate & "=" & strDTEnc & Format(dttmp, "dd/MMM/yyyy") & _
    strDTEnc & " and Cat='" & IIf(GetFlagStatus("LOCATIONWISEHL"), typEmpRot.strLocation, typEmpRot.strCat) & "'", ConMain        ' 18-01
If Not (adrsDept1.EOF And adrsDept1.BOF) Then GetHoliday = pVStar.HlsCode
Exit Function
ERR_P:
    ShowError ("GetHoliday :: Rotation Module")
End Function

Public Sub FixedDaysShiftsOLD(ByVal strEmpCode As String, strMonth As String, _
ByVal strYear As String)    '' Procedure to Assign Shift if Rotation type is 'D'
On Error GoTo ERR_P
Dim dttmp As Date, strTmp As String
Call SetShiftsAll       '' Set the Shifts Array to Blank
typPA.bytFD = 0
typPA.bytShift = 0
bytDaysLeft = 0
'' Move The shift pointer till the Patterns Completed
If typSENum.bytStart = 1 Then Call GetPointerToPatternArray(strMonth, strYear)
If typSENum.bytStart < Val(strPattern(0)) Then
    For typSENum.bytStart = typSENum.bytStart To CByte(strPattern(0)) - 1
        strShiftsAll(typSENum.bytStart) = strShifts(typPA.bytShift)
        dttmp = DateCompDate(GetDateOfDay(typSENum.bytStart, strMonth, strYear))
        strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
        If strTmp <> "" Then
            strShiftsAll(typSENum.bytStart) = pVStar.WosCode   '' For Week Off
        End If
        strTmp = GetHoliday(dttmp)
        If strTmp <> "" Then
            strShiftsAll(typSENum.bytStart) = strTmp  '' For Holiday
        End If
    Next
End If
typPA.bytShift = 0
strShiftsAll(typSENum.bytStart) = strShifts(typPA.bytShift)     '' Put the Next Shift
    dttmp = DateCompDate(GetDateOfDay(typSENum.bytStart, strMonth, strYear))

strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
If strTmp <> "" Then
    strShiftsAll(typSENum.bytStart) = pVStar.WosCode   '' For Week Off
End If
strTmp = GetHoliday(dttmp)
If strTmp <> "" Then
    strShiftsAll(typSENum.bytStart) = strTmp  '' For Holiday
End If
typSENum.bytStart = typSENum.bytStart + 1       '' Increment the Counter
typPA.bytFD = typPA.bytFD + 1                             '' Move Patterns Pointer
If typPA.bytFD > UBound(strPattern) Then typPA.bytFD = 0  '' Reset if Exceeds
Do
    dttmp = DateCompDate(GetDateOfDay(typSENum.bytStart, strMonth, strYear))
    If typSENum.bytStart = CByte(strPattern(typPA.bytFD)) Then    '' If Next Fixed Day
        typPA.bytFD = typPA.bytFD + 1                             '' Move Patterns Pointer
        If typPA.bytFD > UBound(strPattern) Then typPA.bytFD = 0  '' Reset if Exceeds
        typPA.bytShift = typPA.bytShift + 1                       '' Move Shifts Pointer
        If typPA.bytShift > UBound(strShifts) Then typPA.bytShift = 0   '' Reset if Exceeds
    End If
    strShiftsAll(typSENum.bytStart) = strShifts(typPA.bytShift)     '' Put the Next Shift
    strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = pVStar.WosCode   '' For Week Off
    End If
    strTmp = GetHoliday(dttmp)
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = strTmp  '' For Holiday
    End If
    typSENum.bytStart = typSENum.bytStart + 1       '' Increment the Counter
Loop While typSENum.bytStart <= typSENum.bytEnd   '' Move to Next Day
Exit Sub
ERR_P:
    ShowError ("FixedDaysShifs :: Rotation Module")
End Sub
''Modification received from dinesh(India) 24-07-03
Public Sub FixedDaysShifts(ByVal strEmpCode As String, strMonth As String, _
ByVal strYear As String)    '' Procedure to Assign Shift if Rotation type is 'D'
On Error GoTo ERR_P
Dim dttmp As Date, strTmp As String
Dim dtLast As Date, dtStart As Date
Call SetShiftsAll       '' Set the Shifts Array to Blank
typPA.bytFD = 0
typPA.bytShift = 0
bytDaysLeft = 0
dttmp = typEmpRot.dtShift
'' Make the Shift Start Date of the Current Month
dtStart = GetDateOfDay(typSENum.bytStart, strMonth, strYear)
'' Make the Shift Start Date of the Current Month
dtLast = FdtLdt(MonthNumber(strMonth), strYear, "L")

    Dim test As Integer
    test = 0
If dtStart <> typEmpRot.dtShift Then
    For test = 0 To UBound(strPattern) Step 1
        If Val(Day(dttmp)) <= Val(CByte(strPattern(test))) Then
            typPA.bytFD = test
            Exit For
        End If
        typPA.bytShift = typPA.bytShift + 1
    Next
Else
    If Month(dttmp) = Month(dtStart) And Year(dttmp) = Year(dtStart) Then
        For test = 0 To UBound(strPattern) Step 1
            If Val(Day(dttmp)) <= Val(CByte(strPattern(test))) Then
                typPA.bytFD = test
                Exit For
            End If
           typPA.bytShift = typPA.bytShift + 1
        Next
    End If
End If

typPA.bytShift = typPA.bytShift - IIf(typPA.bytShift > 0, 1, 0)
Do While dttmp <= dtLast
    
    
    If Day(dttmp) = CByte(strPattern(typPA.bytFD)) Then    '' If Next Fixed Day
        typPA.bytFD = typPA.bytFD + 1                             '' Move Patterns Pointer
        If typPA.bytFD > UBound(strPattern) Then typPA.bytFD = 0  '' Reset if Exceeds
        'If dttmp <> typEmpRot.dtShift Then
            typPA.bytShift = typPA.bytShift + 1                       '' Move Shifts Pointer
            If typPA.bytShift > UBound(strShifts) Then typPA.bytShift = 0   '' Reset if Exceeds
        'End If
    End If
    If dttmp >= dtStart Then
        strShiftsAll(typSENum.bytStart) = strShifts(typPA.bytShift)
        strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
        If strTmp <> "" Then
            strShiftsAll(typSENum.bytStart) = pVStar.WosCode   '' For Week Off
        End If
        strTmp = GetHoliday(dttmp)
        If strTmp <> "" Then
            strShiftsAll(typSENum.bytStart) = strTmp  '' For Holiday
        End If
        typSENum.bytStart = typSENum.bytStart + 1       '' Increment the Counter
    End If
    dttmp = dttmp + 1
Loop
Exit Sub
ERR_P:
    ShowError ("FixedDaysShifs :: Rotation Module")
End Sub

Public Sub SpecificDaysShifts(ByVal strEmpCode As String, ByVal strMonth As String, _
ByVal strYear As String)
On Error GoTo ERR_P
Dim dttmp As Date, strTmp As String
Dim bytCntPat As Byte, bytCntShift As Byte, bytTmp As Byte
typPA.bytShift = 0
Call SetShiftsAll       '' Set the Shifts Array to Blank
bytTmp = typSENum.bytStart
If typSENum.bytStart = 1 Then bytCntPat = GetPointerToPatternArray(strMonth, strYear)
'' Get the Value of bytCntShift i.e the Array of Pattern

Do While bytTmp <= typSENum.bytEnd
    For bytCntPat = bytCntPat To UBound(strPattern)     '' For Elements in Skip Pattern
        ''typPA.bytShift = bytcntpat
        If typPA.bytShift > UBound(strShifts) Then typPA.bytShift = 0
        For bytCntShift = 1 To CByte(strPattern(bytCntPat)) - bytDaysLeft '' For Elements in Shifts
            strShiftsAll(bytTmp) = strShifts(typPA.bytShift)
            bytTmp = bytTmp + 1                 '' For Next Elements in Sll Shifts Array
            If bytTmp > typSENum.bytEnd Then
                bytCntShift = CByte(strPattern(bytCntPat)) + 2  '' Exit All Loop
                bytCntPat = UBound(strPattern) + 1
            End If
        Next
        bytDaysLeft = 0
        typPA.bytShift = typPA.bytShift + 1
    Next
    bytCntPat = 0
Loop
Do While typSENum.bytStart <= typSENum.bytEnd   '' Move to Next Day
    dttmp = DateCompDate(GetDateOfDay(typSENum.bytStart, strMonth, strYear))
    strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = pVStar.WosCode  '' For Week Off
    End If
    strTmp = GetHoliday(dttmp)
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = strTmp  '' For Holiday
    End If
    typSENum.bytStart = typSENum.bytStart + 1       '' Increments the Counter
Loop
Exit Sub
ERR_P:
    ShowError ("SpecificDaysShifts :: Rotation Module")
End Sub

Private Function GetPointerToPatternArray(ByVal strMonth As String, _
ByVal strYear As String) As Byte
On Error Resume Next
Dim lngTmp As Long, bytCnt As Byte
lngTmp = MonthNumber(strMonth) - 1
If lngTmp = 0 Then
    strYear = CStr(CInt(strYear) - 1)
    lngTmp = 12
End If
''Takeing no of days difference from the shift date to last month's end.
''so that we can determine the occurance of a particular pattern till date.
lngTmp = DateDiff("d", typEmpRot.dtShift, FdtLdt(lngTmp, strYear, "L"))
lngTmp = (lngTmp + 1) Mod bytTotalSkipDays
'' Special Conditions
If lngTmp = 0 Then
    typPA.bytShift = 0
    Exit Function
End If
'' Function to find a particular number of days Over in the last pattern
bytCnt = 0
Do While True
    If bytCnt > UBound(strPattern) Then bytCnt = LBound(strPattern)
    If typPA.bytShift > UBound(strShifts) Then typPA.bytShift = 0
    If lngTmp >= Val(strPattern(bytCnt)) Then
        lngTmp = lngTmp - Val(strPattern(bytCnt))
        bytCnt = bytCnt + 1
        typPA.bytShift = typPA.bytShift + 1
    Else
        GetPointerToPatternArray = bytCnt
        bytDaysLeft = lngTmp
        Exit Do
    End If
Loop
End Function

Public Sub FixedShifts(ByVal strEmpCode As String, ByVal strMonth As String, _
ByVal strYear As String)    '' Procedure to Assign Shift if Shift type is Fixed
On Error GoTo ERR_P
Dim dttmp As Date, strTmp As String
Call SetShiftsAll       '' Set the Shifts Array to Blank
Do While typSENum.bytStart <= typSENum.bytEnd
    frmShiftCr.Refresh
    strShiftsAll(typSENum.bytStart) = typEmpRot.strShiftCode
    dttmp = DateCompDate(GetDateOfDay(typSENum.bytStart, strMonth, strYear))
    strTmp = GetWeekOff(GetWeekDayName(dttmp, 2))
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = pVStar.WosCode  '' For Week Off
    End If
    strTmp = GetHoliday(dttmp)
    If strTmp <> "" Then
        strShiftsAll(typSENum.bytStart) = strTmp  '' For Holiday
    End If
    typSENum.bytStart = typSENum.bytStart + 1       '' Increment the Counter
Loop
Exit Sub
ERR_P:
    ShowError ("FixedShifts :: Rotation Module")
End Sub

Public Sub AdjustSENums(dttmp As Date)
'' dttmp is for Employee Shift Date
typSENum.bytStart = Day(dttmp)
End Sub

Public Sub AddRecordsToShift(ByVal strMonth As String, ByVal strYear As String, _
ByVal strEmpCode As String) '' Procedure to Add Shifts if Monthly Shift File
On Error GoTo ERR_P
Dim bytTmp As Byte, strTmp As String
strTmp = strEmpCode & "'"


For bytTmp = 1 To 31
    strTmp = strTmp & ",'" & strShiftsAll(bytTmp) & "'"
Next
'' Insert Record
ConMain.Execute "Delete From " & Left(strMonth, 3) & _
Right(strYear, 2) & "Shf Where EmpCode='" & strEmpCode & "'"
ConMain.Execute "insert into " & Left(strMonth, 3) & _
Right(strYear, 2) & "Shf Values('" & strTmp & ")"

Exit Sub
ERR_P:
    ShowError ("AddRecordsToShift :: Rotation Module")
End Sub

Public Sub SetDeclareHoliday(ByVal Shmonth As String, ByVal ShYear As String)
    Dim rsgen As New ADODB.Recordset
    Dim Firstdt As Date
    Dim Lastdt As Date
    Dim WeekOff As Boolean
    Dim Qry As String
    Firstdt = "01" & "/" & Shmonth & "/" & ShYear
    Lastdt = DateAdd("d", -1, CDate("01" & "/" & MonthName(MonthNumber(Shmonth) + 1) & "/" & ShYear))
    Qry = "Select * from DeclWoHl Where Date Between " & strDTEnc & Format(Firstdt, "dd/MMM/yy") & strDTEnc & " and " & strDTEnc & Format(Lastdt, "dd/MMM/yy") & strDTEnc & " "
    rsgen.Open Qry, ConMain, adOpenDynamic, adLockOptimistic
    If rsgen.RecordCount < 1 Then Exit Sub
    For i = 1 To rsgen.RecordCount
        If rsgen.Fields("declas") = "WO" Then
            WeekOff = True
        Else
            WeekOff = False
        End If
        Call SetShift(rsgen.Fields("Date"), rsgen.Fields("cat"), rsgen.Fields("compensdt"), WeekOff)
        rsgen.MoveNext
    Next
    
End Sub

Private Sub SetShift(ByVal strDateText As String, ByVal strCatText As String, ByVal strCompDate As String, OptWeek As Boolean)
On Error GoTo ERR_P
Dim adrsEmpCnt As New ADODB.Recordset
Dim strMonShf As String, strTempShfFile As String, strTempShift As String, Ns As String
''
strMonShf = Left(MonthName(Month(DateCompDate(strDateText))), 3)
strMonShf = strMonShf & Right(Year(DateCompDate(strDateText)), 2) & "shf"
''
strTempShfFile = Left(MonthName(Month(DateCompDate(strCompDate))), 3)
strTempShfFile = strTempShfFile & Right(Year(DateCompDate(strCompDate)), 2) & "shf"
''
If FindTable(strMonShf) Then
    If adrsEmpCnt.State = 1 Then adrsEmpCnt.Close
        adrsEmpCnt.Open "Select Empcode from empmst where cat='" & strCatText & "'" _
        , ConMain
 
    Do While Not adrsEmpCnt.EOF
        Ns = Day(DateCompDate(strDateText))
        If adrsRits.State = 1 Then adrsRits.Close   ' 12-08 can't compensate if both date having shift
        adrsRits.Open "Select d" & Trim(Day(DateCompDate(strDateText))) & ",d" & Trim(Day(DateCompDate(strCompDate))) & " from " & strMonShf & _
        " where Empcode=" & "'" & adrsEmpCnt(0) & "'", ConMain
        If Not (adrsRits.EOF And adrsRits.BOF) Then
            If ((adrsRits(0) <> pVStar.HlsCode And adrsRits(0) <> pVStar.WosCode) And (adrsRits(1) <> pVStar.HlsCode And adrsRits(1) <> pVStar.WosCode)) Or (adrsRits(0) = pVStar.WosCode Or adrsRits(0) = pVStar.HlsCode) Then
                adrsEmpCnt.MoveNext
            Else
                If adrsRits.State = 1 Then adrsRits.Close
                adrsRits.Open "Select d" & Ns & " from " & strMonShf & _
                " where Empcode=" & "'" & adrsEmpCnt(0) & "'", ConMain
                If Not (adrsRits.EOF And adrsRits.BOF) Then
                    strTempShift = adrsRits(0)
                    If OptWeek = True Then
                        ConMain.Execute "Update " & strMonShf & " set d" & Ns & _
                        "='" & pVStar.WosCode & "' where Empcode='" & adrsEmpCnt(0) & "'"
                    Else
                        ConMain.Execute "Update " & strMonShf & " set d" & Ns & _
                        "='" & pVStar.HlsCode & "' where Empcode='" & adrsEmpCnt(0) & "'"
                    End If
                      ConMain.Execute "Update MonthShift set d" & Ns & _
                      "='" & pVStar.WosCode & "' Where EmpCode='" & adrsEmpCnt(0) & "' and Month = '" & Left(MonthName(Month(strDateText)), 3) & "' and Yr = " & pVStar.YearSel & " "
                    If FindTable(strTempShfFile) Then
                        Ns = "D" & Trim(Day(DateCompDate(strCompDate)))
                        ConMain.Execute "update " & strTempShfFile & " set " & _
                        Ns & "=" & "'" & strTempShift & "'" & _
                        " where Empcode=" & "'" & adrsEmpCnt(0) & "'"
                    End If
                End If
                adrsEmpCnt.MoveNext
            End If
        Else
            adrsEmpCnt.MoveNext
        End If
    Loop
    adrsEmpCnt.Close
End If
Exit Sub
ERR_P:
    ShowError ("SetShift :: Set Holiday  " & Err.Description)
End Sub



Public Sub ChangeShiftForRobo(strEmpCode As String, strMonth As String, _
    strYear As String)
    Dim adrsTemp As Recordset
    Dim strData() As Variant
    Dim byteStartEnd As StartEndIndex
    Dim WeekWalker As Byte
    Dim WeekDayWalker As Byte
    Dim StartWeekDay As Byte
    Dim UpToWeekDay As Byte
    Dim MonthEnd As Byte
    Dim dttmp As Date
    Set adrsTemp = OpenRecordSet("SELECT * FROM WeekOFF WHERE code='" & _
    GetDeptCode(strEmpCode) & "'")
    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
        strData = adrsTemp.GetRows
    Else
        Exit Sub
    End If
  
    byteStartEnd = GetStartEndIndex(strMonth, strYear)
    StartWeekDay = byteStartEnd.StartIndex
    UpToWeekDay = 8
    MonthEnd = Val(Day(GetMonthEnd(strMonth, strYear)))
    For WeekWalker = 0 To 5
        'MsgBox WeekWalker
        If WeekWalker = 5 Then
            UpToWeekDay = byteStartEnd.EndIndex - MonthEnd
        End If
        For WeekDayWalker = StartWeekDay To UpToWeekDay
            If strData(WeekDayWalker, WeekWalker) = strChecked Then
                If GetDay(WeekWalker, WeekDayWalker, _
                byteStartEnd.StartIndex, 8) <> 0 Then
                    If strShiftsAll(GetDay(WeekWalker, WeekDayWalker, _
                        byteStartEnd.StartIndex, 8)) <> EmptyString Then
                            strShiftsAll(GetDay(WeekWalker, WeekDayWalker, _
                            byteStartEnd.StartIndex, 8)) = pVStar.WosCode
                    End If
                End If
            End If
        Next
        StartWeekDay = 2
        UpToWeekDay = 8
    Next
End Sub
Private Function GetDay(Week As Byte, WeekDay As Byte, StartWeek As Byte, _
    EndWeek As Byte) As Byte
    GetDay = (EndWeek - StartWeek) + ((Week - 1) * 7) + (WeekDay)
    If GetDay > 31 Then
        GetDay = 0
    End If
End Function

Private Function GetStartEndIndex(strMonth As String, _
    strYear As String) As StartEndIndex
    ' Get the Index of the First day
    Select Case UCase(Left(WeekdayName(WeekDay(Year_Start( _
        CByte(MonthNumber((strMonth))), CInt(strYear)), _
        vbUseSystemDayOfWeek)), 3))
        Case "MON"
            GetStartEndIndex.StartIndex = WeekDays.mon
        Case "TUE"
            GetStartEndIndex.StartIndex = WeekDays.Tue
        Case "WED"
            GetStartEndIndex.StartIndex = WeekDays.wed
        Case "THU"
            GetStartEndIndex.StartIndex = WeekDays.thu
        Case "FRI"
            GetStartEndIndex.StartIndex = WeekDays.fri
        Case "SAT"
            GetStartEndIndex.StartIndex = WeekDays.sat
        Case "SUN"
            GetStartEndIndex.StartIndex = WeekDays.sun
    End Select
    GetStartEndIndex.EndIndex = Day(GetMonthEnd(strMonth, strYear)) + _
    GetStartEndIndex.StartIndex
End Function


Private Function GetDeptCode(strEmpCode As String)
    GetDeptCode = ExecScalar("SELECT WeekOffRule FROM empmst WHERE empcode='" & _
    strEmpCode & "'")
End Function


Private Sub SetShiftsAll()          '' Sets the Shifts Array to Blank
    Erase strShiftsAll
End Sub

Private Function AlternateWO(ByVal bytAlt As Byte) As Boolean
Select Case bytAlt
    Case 1      '' Odd
        Select Case typSENum.bytStart
            Case Is <= 7
                AlternateWO = True
            Case 15 To 21
                AlternateWO = True
            Case 29, 30, 31
                AlternateWO = True
            Case Else
                AlternateWO = False
        End Select
    Case 2      '' Even
        Select Case typSENum.bytStart
            Case 8 To 14
                AlternateWO = True
            Case 22 To 28
                AlternateWO = True
            Case Else
                AlternateWO = False
        End Select
End Select
End Function

Public Sub UpdateAfterShiftDate(ByVal strMonth As String, ByVal strYear As String, _
ByVal strEmpCode As String) '' Procedure to Update Shifts if Monthly Shift File
On Error GoTo ERR_P
Dim bytTmp As Byte, strTmp As String        '' If Same Month as Shift Date then process
bytTmp = 1
If UCase(MonthName(Month(adrsEmp("Shf_Date")))) = UCase(strMonth) Then  '' fromDay
    If Year(DateCompDate(adrsEmp("Shf_Date"))) = CInt(strYear) Then _
    bytTmp = Day(adrsEmp("Shf_date"))
End If
strTmp = ""
For bytTmp = bytTmp To typSENum.bytEnd
    strTmp = strTmp & "D" & bytTmp & "='" & strShiftsAll(bytTmp) & "',"
Next
If Right(strTmp, 1) = "," Then strTmp = Left(strTmp, Len(strTmp) - 1)
ConMain.Execute "Update " & Left(strMonth, 3) & _
Right(strYear, 2) & "Shf Set " & strTmp & " Where EmpCode='" & strEmpCode & "'"
Exit Sub
ERR_P:
    ShowError ("Update After Shift Date :: Rotation Module")
End Sub
