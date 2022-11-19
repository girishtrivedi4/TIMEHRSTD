Attribute VB_Name = "mdlDaily"
'' Daily Processing Module
'' -----------------------
Option Explicit
'' Other Variables
Dim VarArrPunches() As Variant    '' Array of Valid Punches for two Days
Public typVar As New clsVars
Public typPerm As New clsPerm
Public typEmp As New clsEmp
Public typCat As New clsCat
Public typShift As New clsShift
Public typDT As New clsDaily
Public typTR As New clsTimes
Public typDH As New clsHours
Public typBH As New clsBreak
Public typOTVars As New clsOTVars
Public typCOVars As New clsCOVars
Dim blnCurrDtTrnFound As Boolean
Dim arrText As String
Dim fs1, f1

Public strArr As String, strdep As String, dte1 As String
Public strarr1 As Single, strdep1 As Single, strarr2 As Single, strdep2 As Single
Public bytIN As Integer, bytOUT As Integer
Dim strArrIO() As String
Public blnIrregular As Boolean
Dim StatusChangLEHour As Boolean, AdFieldChangeSt As Boolean      ' 02-06-09
''
Dim strTotalDays As Double
Dim strLeaveCode As String
Dim sngLeaveAccu As Single
Dim strFlag As String, tempshift As String, newshift As String
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public blnSP As Boolean

Public Function AppendDataFile(frm As Form) As Boolean  '' Function to Append
On Error GoTo Err_particular                                    '' Data from Dat File
Dim strDataL As String                                          '' to the Database
Dim strTmp As String, strListCap As String
Dim j As Integer
strTmp = frm.Caption
AppendDataFile = True
Call TruncateTable("Tbldata")
Dim MDBCON As New ADODB.Connection
Dim dbpath() As String
    Dim adrsdat As New ADODB.Recordset
    adrsdat.Open "Select DatPath From Install", ConMain
    If Len(Trim(adrsdat.Fields(0).Value)) < 1 Then
        MsgBox "Attendance Capturing Dabase File Is Not Selected"
        AppendDataFile = False
        Exit Function
    Else
        dbpath = Split(adrsdat.Fields(0).Value, "|")
        AppendDataFile = True
    End If
    
    For j = 0 To UBound(dbpath)

        MDBCON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath(j) & ";Persist Security Info=False"
        MDBCON.Open
        MDBCON.CursorLocation = adUseClient
        
        If InStr(1, dbpath(j), "eTime", vbTextCompare) Then
            Dim k As Integer
           Dim EmpFrmt As String
        
            Select Case pVStar.CodeSize
                Case 3: EmpFrmt = "000"
                Case 4: EmpFrmt = "0000"
                Case 5: EmpFrmt = "00000"
                Case 6: EmpFrmt = "000000"
            End Select
            If adrsdat.State = 1 Then adrsdat.Close
            For k = Month(typDT.dtFrom) To Month(typDT.dtTo)
                strTmp = "DeviceLogs_" & k & "_" & Year(typDT.dtFrom)
                
                strDataL = "SELECT LogDate, format(UserId, '" & EmpFrmt & "') + C1 FROM " & strTmp & " WHERE LogDate> #" & Format(typDT.dtFrom, "dd/MMM/yy") & "# AND LogDate<#" & Format(typDT.dtTo, "dd/MMM/yy") & "#+1"
                adrsdat.Open strDataL, MDBCON, adLockOptimistic, adLockOptimistic
                Dim adrstbl As New ADODB.Recordset
                adrstbl.Open "Select * from tblData", ConMain, adLockOptimistic, adLockOptimistic
                If adrsdat.RecordCount > 0 Then
                    For i = 0 To adrsdat.RecordCount - 1
                        adrstbl.AddNew
                        adrstbl.Fields(0) = adrsdat.Fields(0)
                        adrstbl.Fields(1) = Replace(Replace(adrsdat.Fields(1), "in", "I"), "out", "O")
                        adrsdat.MoveNext
                    Next
                    adrstbl.UpdateBatch
                End If
                adrstbl.Close
                adrsdat.Close
            Next
        
        Else
    
            If adrsdat.State = 1 Then adrsdat.Close
            
            If typPerm.blnIO = False Then
            
                strDataL = "SELECT CHECKTIME ,   USERINFO.Badgenumber  as c  FROM CHECKINOUT INNER JOIN USERINFO ON " & _
                 " CHECKINOUT.USERID = USERINFO.USERID " & _
                " Where CHECKTIME > #" & Format(typDT.dtFrom, "dd/MMM/yyyy") & "# AND CHECKTIME< #" & Format(typDT.dtTo, "dd/mmm/yyyy") & "# +1"
            Else
                strDataL = "SELECT CHECKTIME ,   USERINFO.Badgenumber + CHECKINOUT.CHECKTYPE as c  FROM CHECKINOUT INNER JOIN USERINFO ON " & _
                 " CHECKINOUT.USERID = USERINFO.USERID " & _
                " Where CHECKTIME > #" & Format(typDT.dtFrom, "dd/MMM/yyyy") & "# AND CHECKTIME< #" & Format(typDT.dtTo, "dd/mmm/yyyy") & "# +1"
            End If
            
            adrsdat.Open strDataL, MDBCON
            
            Dim adrstbl1 As New ADODB.Recordset
            adrstbl1.Open "Select * from tblData", ConMain, adOpenDynamic, adLockOptimistic
            
            If adrsdat.RecordCount > 0 Then
            
                For i = 0 To adrsdat.RecordCount - 1
                    adrstbl1.AddNew
                    adrstbl1.Fields(0) = adrsdat.Fields(0)
                    adrstbl1.Fields(1) = adrsdat.Fields(1)
                    adrsdat.MoveNext
                Next
                adrstbl1.UpdateBatch
            End If
            MDBCON.Close
            adrstbl1.Close
        End If
    Next
    
    If GetFlagStatus("DeviceLog") Then
        If typPerm.blnIO = False Then
            If bytBackEnd = "2" Then
                sqlStr = " INSERT INTO tblData ( strF1, strCode ) SELECT DeviceLog.PDate, DeviceLog.CardCode FROM DeviceLog WHERE (((DeviceLog.[PDate])> #" & Format(typDT.dtFrom, "dd/MMM/yyyy") & "#)  AND (DeviceLog.[PDate])< #" & Format(typDT.dtTo, "dd/mmm/yyyy") & "# +1)"
            Else
                arrText = "DateAdd(Day,1," & strDTEnc & Format(typDT.dtTo, "dd/mmm/yyyy") & strDTEnc & ")"
                sqlStr = " INSERT INTO tblData ( strF1, strCode ) SELECT DeviceLog.PDate, DeviceLog.CardCode FROM DeviceLog WHERE (((DeviceLog.[PDate])> " & strDTEnc & Format(typDT.dtFrom, "dd/MMM/yyyy") & strDTEnc & ")  AND (DeviceLog.[PDate])<  " & arrText & " )"
            End If
        Else
            If bytBackEnd = "2" Then
                sqlStr = " INSERT INTO tblData ( strF1, strCode ) SELECT DeviceLog.PDate, DeviceLog.CardCode FROM DeviceLog WHERE (((DeviceLog.[PDate])> #" & Format(typDT.dtFrom, "dd/MMM/yyyy") & "#)  AND (DeviceLog.[PDate])< #" & Format(typDT.dtTo, "dd/mmm/yyyy") & "# +1)"
            Else
                arrText = "DateAdd(Day,1," & strDTEnc & Format(typDT.dtTo, "dd/mmm/yyyy") & strDTEnc & ")"
                sqlStr = " INSERT INTO tblData ( strF1, strCode ) SELECT DeviceLog.PDate, DeviceLog.CardCode FROM DeviceLog WHERE (((DeviceLog.[PDate])> " & strDTEnc & Format(typDT.dtFrom, "dd/MMM/yyyy") & strDTEnc & ")  AND (DeviceLog.[PDate])< " & arrText & ")"
            End If
        End If
        ConMain.Execute sqlStr
    End If

frm.Caption = strTmp
Exit Function
Err_particular:
    Call ShowError("Capturing Database File Error" & _
    vbCrLf & "Please Retry")
    AppendDataFile = False
End Function

Private Function Decrypt(mData As String)        ' 03-02
Dim k As Long
Dim mChr As String
Dim mAsc As Long
Dim mSeed As Long
If mData = "" Then Exit Function
mData = Mid(mData, 1, Len(mData) - 1)
mSeed = (Asc(Right(mData, 1)) - 9) / 9
 For k = 1 To Len(mData) - 1
    mAsc = Asc(Mid(mData, k, 1)) - mSeed
    If mAsc < 0 Then mAsc = 255 + mAsc
    Decrypt = Decrypt & Chr(mAsc)
 Next k
End Function



Private Function GetDataFromSQL(strFileN As String, strPath As String)
    ConMain.Execute "INSERT INTO Tbldata " & _
    " SELECT A.* FROM OPENROWSET('MSDASQL', " & _
    " 'Driver={Microsoft Text Driver (*.txt; *.csv)}; " & _
    "Dbq=" & strPath & ";Extensions=asc,csv,tab,txt', '" & _
    " SELECT * FROM [" & strFileN & "]') AS A WHERE (strF1 IS NOT NULL or " & _
    " strf1<>'')"
End Function

Private Function WriteToSchema(strMainKey As String, strPath As String)
On Error GoTo Err
    Call WriteINIString(strMainKey, "ColNameHeader", _
     "False", strPath & "\schema.ini")
    Call WriteINIString(strMainKey, "Format", _
     "FixedLength", strPath & "\schema.ini")
    Call WriteINIString(strMainKey, "CharacterSet", _
     "ANSI", strPath & "\schema.ini")
    Call WriteINIString(strMainKey, "Col1", _
     "STRF1 TEXT", strPath & "\schema.ini")
Exit Function
Err:
    Call ShowError("Error in WriteToSchema")
End Function

Public Function ChangeExtension(ByVal FolderName As String, _
  ByVal NewExtension As String, strFileName As String, _
  Optional ByVal OldExtension As _
  String = "") As Boolean
    Dim oFso As New FileSystemObject
    Dim oFolder As Folder
    Dim oFile As File
    Dim sOldName As String
    Dim sNewName As String
    Dim iCtr As Long
    Dim iDotPosition As Integer
    Dim sWithoutExt As String
    Dim sFolderName As String
    
    sFolderName = FolderName
    If Right(sFolderName, 1) <> "\" Then sFolderName = _
       sFolderName & "\"
    Set oFolder = oFso.GetFolder(FolderName)
    sOldName = sFolderName & strFileName
    sNewName = ""
    iDotPosition = InStrRev(sOldName, ".")
    If iDotPosition > 0 Then
        If OldExtension = "" Or UCase(Mid(sOldName, _
           iDotPosition + 1)) = UCase(OldExtension) Then
                sWithoutExt = Left(sOldName, iDotPosition - 1)
                sNewName = sWithoutExt & "." & NewExtension
                Name sOldName As sNewName
                Err.clear
                On Error GoTo errorHandler
         End If
    End If
    ChangeExtension = True
    Set oFile = Nothing
    Set oFolder = Nothing
    Set oFso = Nothing
Exit Function
errorHandler:
    Call ShowError("ChangeExtension ")
End Function

Private Function CheckShiftPunches() As Boolean      '' Function to Check Existence of Shift
On Error GoTo Err_particular                        '' and Punches
Dim bytAction As Integer
Dim blnCSP As Boolean
Dim dtShift As Date
bytAction = 0
typVar.strShiftTmp = ""
typVar.strShiftOfDay = ""
Call GetShiftOfDay(CStr(typDT.dtFrom), typEmp.strEmp)
Select Case typVar.strShiftOfDay
    '' Weekoff or Holiday
    Case typVar.strWosCode, typVar.strHlsCode
        typVar.strShiftTmp = typVar.strShiftOfDay
        typVar.strShiftOfDay = ""
        Call WOHLAction
    '' Blank Shift
    Case ""
        Call ActionBlank
    '' Other Shift
    Case Else
        Call FillShifttype(typVar.strShiftOfDay)
        '' If Invalid Shift is Found
        If typShift.sngIN = 0 Then
            typVar.strShiftOfDay = ""
            Call ActionBlank
        End If
End Select
'' New Code Till Here

If typVar.strShiftOfDay = "" And (typVar.strShiftTmp <> pVStar.WosCode Or Not typEmp.blnAutoOnPunch) Then
''
    bytAction = 1       '' No Shift is Found
Else
    '' If WO or Holiday Put Status Accordingly
    Select Case ValidatePunches
        Case 0      '' Punches With Zero Flag Found
            bytAction = 2
        Case 1      '' No Punches Found
            bytAction = 3
        Case 2      '' No Punches With Zero Flag Found
            bytAction = 4
    End Select
End If
Select Case bytAction
    Case 1  '' No Shift is Found
        Call NoShift
        blnCSP = False
    Case 2  '' All Shifts and Punches are Found
        
        If (typPerm.blnIO Or typPerm.blnDI) Then
            Call PutTimeRangeIO
        Else
            Call PutTimeRange
        End If
        ''
        Call DeleteUsedPunches
        blnCSP = True
    Case 3  '' No Punches are Found
        GetStatus (1)    '' Get the Normal Status First
        Call SettypTR
        Call SettypDH
        '' Set OT & CO Vars
        Call SettypOTVars
        Call SettypCOVars
        typVar.sngPresent = 1
        Call AddRecordsToTrn
        blnCSP = False
    Case 4  '' No Punches With Zero Flags are Found
        GetStatus (1)     '' Get the Normal Status First
        Call SettypTR
        Call SettypDH
        '' Set OT & CO Vars
        Call SettypOTVars
        Call SettypCOVars
        typVar.sngPresent = 1
        Call AddRecordsToTrn
        blnCSP = False
End Select
CheckShiftPunches = blnCSP
tempshift = typVar.strShiftOfDay
Exit Function
Err_particular:
    Call ShowError("CheckShiftPunches")
End Function

Private Sub WOHLAction()
On Error GoTo ERR_P
Dim strquery As String, strTmp As String, dttmp As Date, blnTmp As Boolean
Dim adrsD As New ADODB.Recordset
With adrsD
    If .State = 1 Then .Close
    .ActiveConnection = ConMain
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
End With
Select Case typEmp.bytWOHLAction
    Case 0
        dttmp = typDT.dtFrom - 1
        strTmp = Left(MonthName(Month(dttmp)), 3) & Right(Year(dttmp), 2) & "Trn"
        strquery = "Shift"
    Case 1
        dttmp = typDT.dtFrom + 1
        strTmp = Left(MonthName(Month(dttmp)), 3) & Right(Year(dttmp), 2) & "Shf"
        strquery = "D" & Day(dttmp)
End Select
'' If Assign Previous Day Shift
If typEmp.bytWOHLAction = 0 Then
    If FindTable(strTmp) Then
        adrsD.Open "Select " & strquery & " from " & strTmp & " Where Empcode='" & _
        typEmp.strEmp & "' and " & strKDate & "=" & strDTEnc & Format(dttmp, "dd/mmm/yyyy") & _
        strDTEnc
        If Not adrsD.EOF Then
            If Not IsNull(adrsD(strquery)) Then
                Select Case adrsD(strquery)
                    Case ""
                        '' Blank Shift Found
                        blnTmp = True
                    Case Else
                        Call FillShifttype(adrsD(strquery))
                        If typShift.sngIN = 0 Then
                            '' If Invalid Shift Found
                            blnTmp = True
                        Else
                            typVar.strShiftOfDay = adrsD(strquery)
                        End If
                End Select
            Else
                '' NULL value found
                blnTmp = True
            End If
        Else
            '' Record not Found
            blnTmp = True
        End If
    Else
        '' File not Found
        blnTmp = True
    End If
End If
'' If Assign Next Day Shift
If typEmp.bytWOHLAction = 1 Then
    If FindTable(strTmp) Then
        adrsD.Open "Select " & strquery & " from " & strTmp & " Where Empcode='" & _
        typEmp.strEmp & "'"
        If Not adrsD.EOF Then
            If Not IsNull(adrsD(strquery)) Then
                Select Case adrsD(strquery)
                    Case "", typVar.strWosCode, typVar.strHlsCode
                        '' Blank Shift,Weekoff,Holiday Found
                        blnTmp = True
                    Case Else
                        Call FillShifttype(adrsD(strquery))
                        If typShift.sngIN = 0 Then
                            '' If Invalid Shift Found
                            blnTmp = True
                        Else
                            typVar.strShiftOfDay = adrsD(strquery)
                        End If
                End Select
            Else
                '' NULL value found
                blnTmp = True
            End If
        Else
            '' Record not Found
            blnTmp = True
        End If
    Else
        '' File not Found
        blnTmp = True
    End If
End If
'' Assign a Specific Shift
If typEmp.bytWOHLAction = 2 Then
    If typEmp.strAction3Shift <> "" Then
        Call FillShifttype(typEmp.strAction3Shift)
        If typShift.sngIN = 0 Then
            '' If Invalid Shift Found
            blnTmp = True
        Else
            typVar.strShiftOfDay = typEmp.strAction3Shift
        End If
    Else
        '' If Blank Shift Assigned
        blnTmp = True
    End If
End If
'' If Blank Shift Found
If blnTmp Then
    Call ActionBlank
End If
If typEmp.blnAutoOnPunch Then typVar.strShiftOfDay = typVar.strShiftTmp
Exit Sub
ERR_P:
    ShowError ("WOHLAction::mdlDaily")
End Sub

Private Sub ActionBlank()
On Error GoTo ERR_P
'' If Keep Blank
If typEmp.strActionBlank = "" Then Exit Sub
Call FillShifttype(typEmp.strActionBlank)
If typShift.sngIN = 0 Then Exit Sub
typVar.strShiftOfDay = typEmp.strActionBlank
Exit Sub
ERR_P:
    ShowError ("ActionBlank::mdlDaily")
End Sub

Private Sub DeleteUsedPunches()     '' Deletes the Punches which are not to be counted
On Error GoTo Err_particular        '' For the Next Day Processing
Dim bytCntTR As Byte
For bytCntTR = 0 To UBound(VarArrPunches)
    If VarArrPunches(bytCntTR, 3) = 0 Or VarArrPunches(bytCntTR, 3) = 1 Then
        ConMain.Execute "Delete from DailyPro where UnqFld=" & _
        VarArrPunches(bytCntTR, 2)
    End If
Next
Exit Sub
Err_particular:
    Call ShowError("DeleteUsedPunches")
End Sub

Public Sub FillInstalltypes() '' Fills Install types
On Error GoTo Err_particular
If adRsInstall.State = 1 Then adRsInstall.Close
adRsInstall.Open "Select * from install", ConMain
'' Permission
typPerm.blnIsPerm = IIf(adRsInstall("prsm_cards") = "1", True, False)
typPerm.strBus = IIf(IsNull(adRsInstall("bus_card")), "", adRsInstall("bus_card"))
typPerm.strEarl = IIf(IsNull(adRsInstall("Earl_card")), "", adRsInstall("Earl_card"))
typPerm.strLate = IIf(IsNull(adRsInstall("Late_Card")), "", adRsInstall("Late_Card"))
typPerm.strOD = IIf(IsNull(adRsInstall("OD_Card")), "", adRsInstall("OD_Card"))
typPerm.strTempC = IIf(IsNull(adRsInstall("Tmp_Card")), "", adRsInstall("Tmp_Card"))
'' Miscellaneous
typPerm.sngPostEarl = IIf(IsNull(adRsInstall("PostErl")), 0, adRsInstall("PostErl"))
typPerm.sngPostLt = IIf(IsNull(adRsInstall("PostLt")), 0, adRsInstall("PostLt"))

''typPerm.sngUpto = IIf(IsNull(adRsInstall("Upto")), 0, adRsInstall("Upto"))
''
typPerm.sngFiltTime = IIf(IsNull(adRsInstall("Filt_Time")), 0, adRsInstall("Filt_Time"))
'
typPerm.intLvUpYr = IIf(IsNull(adRsInstall("lvupdtyear")), 0, adRsInstall("lvupdtyear"))
typPerm.intYrFrom = adRsInstall("yearfrom")

typPerm.blnIO = IIf(adRsInstall("IO") = "Y", True, False)
typPerm.blnIgnore = IIf(adRsInstall("IgnoreP") = "Y", True, False)
''
Exit Sub
Err_particular:
    Call ShowError("FillInstalltypes")
End Sub

Public Sub FillEmptype(ByVal STRECODE As String)    '' Fills Employee type
On Error GoTo Err_particular
If adrsEmp.RecordCount = 0 Then
    Call SettypEmp
    Exit Sub
End If
adrsEmp.MoveFirst
adrsEmp.Find "Empcode=  '" & STRECODE & "'"
If adrsEmp.EOF Then Exit Sub
typEmp.blnAuto = IIf(adrsEmp("shf_chg") = 0, False, True)
typEmp.bytEntry = IIf(IsNull(adrsEmp("Entry")), 0, adrsEmp("Entry"))
typEmp.strECat = IIf(IsNull(adrsEmp("Cat")), "", adrsEmp("Cat"))
typEmp.strEmp = IIf(IsNull(adrsEmp("Empcode")), "", adrsEmp("Empcode"))
typEmp.strName = IIf(IsNull(adrsEmp("Name")), "", adrsEmp("Name"))
typEmp.strCard = IIf(IsNull(adrsEmp("Card")), "", adrsEmp("Card"))
'typEmp.strConv = IIf(IsNull(adrsEmp("Conv")), "", adrsEmp("Conv"))
typEmp.strOff = IIf(IsNull(adrsEmp("Off")), "", adrsEmp("Off"))
typEmp.strOff2 = IIf(IsNull(adrsEmp("Off2")), "", adrsEmp("Off2"))
typEmp.strWO13 = IIf(IsNull(adrsEmp("WO_1_3")), "", adrsEmp("WO_1_3"))
typEmp.strWO24 = IIf(IsNull(adrsEmp("WO_2_4")), "", adrsEmp("WO_2_4"))
typEmp.strShifttype = IIf(IsNull(adrsEmp("Styp")), "", adrsEmp("Styp"))
typEmp.dtJoin = adrsEmp("JoinDate")
typEmp.dtLeft = IIf(IsNull(adrsEmp("Leavdate")), typDT.dtFrom + 1, adrsEmp("Leavdate"))
Select Case typEmp.strShifttype
    Case ""
        typEmp.strEmpShift = ""
    Case "R"
        typEmp.strEmpShift = IIf(IsNull(adrsEmp("SCode")), "", adrsEmp("SCode"))
    Case "F"
        typEmp.strEmpShift = IIf(IsNull(adrsEmp("F_Shf")), "", adrsEmp("F_Shf"))
End Select
'' OT Rule
typOTVars.bytOTCode = IIf(IsNull(adrsEmp("OTCode")), 100, _
adrsEmp("OTCode"))
'' CO Rule
typCOVars.bytCOCode = IIf(IsNull(adrsEmp("COCode")), 100, _
adrsEmp("COCode"))
'' For Details regarding Daily Processing
typEmp.bytWOHLAction = IIf(IsNull(adrsEmp("WOHLAction")), 0, adrsEmp("WOHLAction"))
typEmp.strAction3Shift = IIf(IsNull(adrsEmp("Action3Shift")), "", adrsEmp("Action3Shift"))
typEmp.blnAutoOnPunch = IIf(adrsEmp("AutoForPunch") = 1, True, False)
typEmp.strActionBlank = IIf(IsNull(adrsEmp("ActionBlank")), "", adrsEmp("ActionBlank"))

typEmp.strAutoGroup = IIf(IsNull(adrsEmp("AutoG")), "", adrsEmp("AutoG"))
blnIrregular = False
''
Exit Sub
Err_particular:
    Call ShowError("FillEmptype")
'    Resume Next
End Sub

Public Sub FillCattype(ByVal strCCode As String)    '' Fills Category type
On Error GoTo Err_particular
If AdrsCat.RecordCount = 0 Then
    Call SettypCat
    Exit Sub
End If
AdrsCat.MoveFirst
AdrsCat.Find "cat='" & strCCode & "'"
If AdrsCat.EOF Then Exit Sub
typCat.sngCutE = IIf(IsNull(AdrsCat("HalfCutEr")), 0, AdrsCat("HalfCutEr"))
typCat.sngCutL = IIf(IsNull(AdrsCat("HalfCutLt")), 0, AdrsCat("HalfCutLt"))
typCat.sngEarl = IIf(IsNull(AdrsCat("Erl_Allow")), 0, AdrsCat("Erl_Allow"))
typCat.sngEarlI = IIf(IsNull(AdrsCat("Erl_Ignore")), 0, AdrsCat("Erl_Ignore"))
typCat.sngLate = IIf(IsNull(AdrsCat("Lt_Allow")), 0, AdrsCat("Lt_Allow"))
typCat.sngLateI = IIf(IsNull(AdrsCat("Lt_Ignore")), 0, AdrsCat("Lt_Ignore"))
typCat.strCat = AdrsCat("Cat")
typCat.strDesc = IIf(IsNull(AdrsCat("desc")), "", AdrsCat("desc"))

Exit Sub
Err_particular:
    Call ShowError("FillCattype")
    'Resume Next
End Sub

Public Sub FillOTType(ByVal bytOTCode As Integer)
On Error GoTo Err_particular
If typOTVars.bytOTCode = 100 Then Exit Sub
If adrsOT.RecordCount = 0 Then
    typOTVars.bytOTCode = 100
    Call SettypOTVars
    Exit Sub
End If
adrsOT.MoveFirst
adrsOT.Find "OTCode=" & bytOTCode
If adrsOT.EOF Then
    typOTVars.bytOTCode = 100
    Call SettypOTVars
    Exit Sub
End If
'' Code to Fill OT Variables
typOTVars.bytOTWD = IIf(adrsOT("OTWD") = 1, 1, 0)
typOTVars.bytOTWO = IIf(adrsOT("OTWO") = 1, 1, 0)
typOTVars.bytOTHL = IIf(adrsOT("OTHL") = 1, 1, 0)
'' OT Rates
typOTVars.sngWDRate = IIf(IsNull(adrsOT("WDRates")), "0.00", Format(adrsOT("WDRates")))
typOTVars.sngWORate = IIf(IsNull(adrsOT("WORates")), "0.00", Format(adrsOT("WORates")))
typOTVars.sngHLRate = IIf(IsNull(adrsOT("HLRates")), "0.00", Format(adrsOT("HLRates")))
'' Authorized by Default
typOTVars.strOTAuth = IIf(IsNull(adrsOT("Authorized")), "", adrsOT("Authorized"))
'' Maximum OT
typOTVars.sngMaxOT = IIf(IsNull(adrsOT("MaxOT")), "0.00", Format(adrsOT("MaxOT")))
'' Late-Early Deductions
typOTVars.bytDedLate = IIf(adrsOT("DedLate") = 1, 1, 0)
typOTVars.bytDedEarl = IIf(adrsOT("DedEarl") = 1, 1, 0)
'' Deductions
typOTVars.sngF1 = IIf(IsNull(adrsOT("From1")), "0.00", Format(adrsOT("From1")))
typOTVars.sngT1 = IIf(IsNull(adrsOT("To1")), "0.00", Format(adrsOT("To1")))
typOTVars.sngD1 = IIf(IsNull(adrsOT("Deduct1")), "0.00", Format(adrsOT("Deduct1")))
typOTVars.bytAll1 = IIf(adrsOT("All1") = 1, 1, 0)
typOTVars.sngF2 = IIf(IsNull(adrsOT("From2")), "0.00", Format(adrsOT("From2")))
typOTVars.sngT2 = IIf(IsNull(adrsOT("To2")), "0.00", Format(adrsOT("To2")))
typOTVars.sngD2 = IIf(IsNull(adrsOT("Deduct2")), "0.00", Format(adrsOT("Deduct2")))
typOTVars.bytAll2 = IIf(adrsOT("All2") = 1, 1, 0)
typOTVars.sngF3 = IIf(IsNull(adrsOT("From3")), "0.00", Format(adrsOT("From3")))
typOTVars.sngT3 = IIf(IsNull(adrsOT("To3")), "0.00", Format(adrsOT("To3")))
typOTVars.sngD3 = IIf(IsNull(adrsOT("Deduct3")), "0.00", Format(adrsOT("Deduct3")))
typOTVars.bytAll3 = IIf(adrsOT("All3") = 1, 1, 0)
typOTVars.sngMoreThan = IIf(IsNull(adrsOT("MoreThan")), "0.00", Format(adrsOT("Morethan")))
typOTVars.sngD4 = IIf(IsNull(adrsOT("Deduct4")), "0.00", Format(adrsOT("Deduct4")))
typOTVars.bytAll4 = IIf(adrsOT("All4") = 1, 1, 0)
typOTVars.bytApplyWO = IIf(adrsOT("WODeduct") = 1, 1, 0)
typOTVars.bytApplyHL = IIf(adrsOT("HLDeduct") = 1, 1, 0)
'' Round Off
typOTVars.sngRF1 = IIf(IsNull(adrsOT("RFrom1")), "0.00", Format(adrsOT("RFrom1")))
typOTVars.sngRT1 = IIf(IsNull(adrsOT("RTo1")), "0.00", Format(adrsOT("RTo1")))
typOTVars.sngR1 = IIf(IsNull(adrsOT("Round1")), "0.00", Format(adrsOT("Round1")))
typOTVars.sngRF2 = IIf(IsNull(adrsOT("RFrom2")), "0.00", Format(adrsOT("RFrom2")))
typOTVars.sngRT2 = IIf(IsNull(adrsOT("RTo2")), "0.00", Format(adrsOT("RTo2")))
typOTVars.sngR2 = IIf(IsNull(adrsOT("Round2")), "0.00", Format(adrsOT("Round2")))
typOTVars.sngRF3 = IIf(IsNull(adrsOT("RFrom3")), "0.00", Format(adrsOT("RFrom3")))
typOTVars.sngRT3 = IIf(IsNull(adrsOT("RTo3")), "0.00", Format(adrsOT("RTo3")))
typOTVars.sngR3 = IIf(IsNull(adrsOT("Round3")), "0.00", Format(adrsOT("Round3")))
typOTVars.sngRF4 = IIf(IsNull(adrsOT("RFrom4")), "0.00", Format(adrsOT("RFrom4")))
typOTVars.sngRT4 = IIf(IsNull(adrsOT("RTo4")), "0.00", Format(adrsOT("RTo4")))
typOTVars.sngR4 = IIf(IsNull(adrsOT("Round4")), "0.00", Format(adrsOT("Round4")))
typOTVars.sngRT5 = IIf(IsNull(adrsOT("RTo5")), "0.00", Format(adrsOT("RTo5")))
typOTVars.sngR5 = IIf(IsNull(adrsOT("Round5")), "0.00", Format(adrsOT("Round5")))
Exit Sub
Err_particular:
    Call ShowError("FillOTType")
End Sub

Public Sub FillCOType(ByVal bytCOCode As Integer)
On Error GoTo Err_particular
If typCOVars.bytCOCode = 100 Then Exit Sub
If adrsCO.EOF Then
    typCOVars.bytCOCode = 100
    Call SettypCOVars
    Exit Sub
End If
adrsCO.MoveFirst
adrsCO.Find "COCode=" & bytCOCode
If adrsCO.EOF Then
    typCOVars.bytCOCode = 100
    Call SettypCOVars
    Exit Sub
End If
'' Code to Fill CO Variables
'' Give CO on
typCOVars.bytCOWD = IIf(adrsCO("COWD") = 1, 1, 0)
typCOVars.bytCOWO = IIf(adrsCO("COWO") = 1, 1, 0)
typCOVars.bytCOHL = IIf(adrsCO("COHL") = 1, 1, 0)
'' Continuous Slab
typCOVars.bytCOAvail = IIf(IsNull(adrsCO("COAvail")), 0, adrsCO("COAvail"))
'' Others
typCOVars.sngWDH = IIf(IsNull(adrsCO("WDH")), "0.00", Format(adrsCO("WDH")))
typCOVars.sngWOH = IIf(IsNull(adrsCO("WOH")), "0.00", Format(adrsCO("WOH")))
typCOVars.sngHLH = IIf(IsNull(adrsCO("HLH")), "0.00", Format(adrsCO("HLH")))
typCOVars.sngWDF = IIf(IsNull(adrsCO("WDF")), "0.00", Format(adrsCO("WDF")))
typCOVars.sngWOF = IIf(IsNull(adrsCO("WOF")), "0.00", Format(adrsCO("WOF")))
typCOVars.sngHLF = IIf(IsNull(adrsCO("HLF")), "0.00", Format(adrsCO("HLF")))
'' Late Early
typCOVars.bytCOLate = IIf(adrsCO("DedLate") = 1, 1, 0)
typCOVars.bytCOEarl = IIf(adrsCO("DedEarl") = 1, 1, 0)
Exit Sub
Err_particular:
    Call ShowError("FillCOType")
End Sub

Public Sub FillShifttype(ByVal strSCode As String)  '' Fills Shift type
On Error GoTo Err_particular
If adRsintshft.EOF Then
    Call SettypShift
End If
If adRsintshft.RecordCount > 0 Then
    adRsintshft.MoveFirst
Else
    Exit Sub
End If
adRsintshft.Find "Shift='" & strSCode & "'"
If adRsintshft.EOF Then Exit Sub
typShift.blnNight = IIf(adRsintshft("Night") = 0, False, True)
typShift.sngB1I = IIf(IsNull(adRsintshft("Rst_In")), 0, adRsintshft("Rst_In"))
typShift.sngB1O = IIf(IsNull(adRsintshft("Rst_Out")), 0, adRsintshft("Rst_Out"))
typShift.sngB2I = IIf(IsNull(adRsintshft("Rst_In_2")), 0, adRsintshft("Rst_In_2"))
typShift.sngB2O = IIf(IsNull(adRsintshft("Rst_Out_2")), 0, adRsintshft("Rst_Out_2"))
typShift.sngB3I = IIf(IsNull(adRsintshft("Rst_In_3")), 0, adRsintshft("Rst_In_3"))
typShift.sngB3O = IIf(IsNull(adRsintshft("Rst_Out_3")), 0, adRsintshft("Rst_Out_3"))
typShift.sngBH1 = IIf(IsNull(adRsintshft("Rst_Brk")), 0, adRsintshft("Rst_Brk"))
typShift.sngBH2 = IIf(IsNull(adRsintshft("Rst_Brk_2")), 0, adRsintshft("Rst_Brk_2"))
typShift.sngBH3 = IIf(IsNull(adRsintshft("Rst_Brk_3")), 0, adRsintshft("Rst_Brk_3"))
typShift.sngHalfE = IIf(IsNull(adRsintshft("HDEnd")), 0, adRsintshft("HDEnd"))
typShift.sngHalfS = IIf(IsNull(adRsintshft("HDStart")), 0, adRsintshft("HDStart"))
typShift.sngHRS = IIf(IsNull(adRsintshft("Shf_Hrs")), 0, adRsintshft("Shf_Hrs"))
typShift.sngIN = IIf(IsNull(adRsintshft("Shf_In")), 0, adRsintshft("Shf_In"))
typShift.sngOut = IIf(IsNull(adRsintshft("Shf_Out")), 0, adRsintshft("Shf_Out"))
typShift.strShift = IIf(IsNull(adRsintshft("Shift")), "", adRsintshft("Shift"))

typShift.sngUPTO = IIf(IsNull(adRsintshft("UPTO")), 0, adRsintshft("UPTO"))
''
Exit Sub
Err_particular:
    Call ShowError("FillShifttype")
End Sub

Public Sub FilterOnCard()    '' Filters RAW Data on Employee Card Basis
'On Error GoTo ERR_Particular
Dim strTmp() As String, intTmpFld As Long, intTmpCnt As Long
Dim adrsDumb As New ADODB.Recordset, strDelStr As String
intTmpCnt = -1
'' Filter 2 Level (Permission Card Entries of the Non Selected Employees)
If adrsDumb.State = 1 Then adrsDumb.Close
adrsDumb.Open "Select Card,UnqFld from DailyPro", ConMain
Do While Not adrsDumb.EOF
Start_Loop:
    If adrsDumb("Card") = typPerm.strEarl Or adrsDumb("Card") = typPerm.strLate Or _
    adrsDumb("Card") = typPerm.strOD Then
        intTmpFld = adrsDumb("UnqFld")
    Else
        intTmpFld = -1
    End If
    adrsDumb.MoveNext
    If adrsDumb.EOF Then
        If intTmpFld <> -1 Then
            intTmpCnt = intTmpCnt + 1
            ReDim Preserve strTmp(intTmpCnt)
            strTmp(intTmpCnt) = intTmpFld
        End If
    Else
        If intTmpFld <> -1 Then
            If adrsDumb("UnqFld") - intTmpFld > 1 Then
                intTmpCnt = intTmpCnt + 1
                ReDim Preserve strTmp(intTmpCnt)
                strTmp(intTmpCnt) = intTmpFld
                '' Added on 08-03-2003
''************* FOLLOWING ELSEIF BLOCK IS TO BE ADDED IN ALL THE COPIES.'*************
            ElseIf adrsDumb("UnqFld") - intTmpFld = 1 Then
                If adrsDumb("Card") = typPerm.strEarl Or adrsDumb("Card") = typPerm.strLate Or _
                adrsDumb("Card") = typPerm.strOD Then
                    intTmpCnt = intTmpCnt + 1
                    ReDim Preserve strTmp(intTmpCnt)
                    strTmp(intTmpCnt) = intTmpFld
                Else
                    GoTo Start_Loop
                End If
''*****************************************************************************************
            Else
                GoTo Start_Loop
            End If
        Else
            GoTo Start_Loop
        End If
    End If
Loop
If intTmpCnt = -1 Then Exit Sub
For intTmpCnt = 0 To UBound(strTmp)
    strDelStr = strDelStr & strTmp(intTmpCnt) & ","
Next
strDelStr = Left(strDelStr, Len(strDelStr) - 1)
ConMain.Execute "Delete from DailyPro Where UnqFld in (" & _
strDelStr & ")"
Exit Sub
Err_particular:
    Call ShowError("FilterOnCard")
End Sub

Public Sub FilterOnDates()  '' Filters Data on the Selection Date Basis
On Error GoTo Err_particular
Dim strD1 As Date, strD2 As Date
strD1 = typDT.dtFrom - 1 ' Month(typDT.dtFrom - 1) & Format(Day(typDT.dtFrom - 1), "00")
strD2 = typDT.dtTo + 1 ' Month(typDT.dtTo + 1) & Format(Day(typDT.dtTo + 2), "00")
   
Select Case bytBackEnd
    Case 1
        ConMain.Execute "Delete from tbldata where strf1 Not Between '" & Format(strD1, "dd/mmm/yyyy") & "' And '" & Format(strD2, "dd/mmm/yyyy") & "'" & _
        "" & _
        ""
    Case 2
        ConMain.Execute "Delete from tbldata where strf1 Not Between #" & Format(strD1, "dd/mmm/yyyy") & "# And #" & Format(strD2, "dd/mmm/yyyy") & "#"
    Case 3  ''Oracle
        ConMain.Execute "Delete from tbldata where TO_NUMBER(substr(strF1,3,2) " & _
        "|| substr(strF1,1,2))" & "" & strD1 & " and " & strD2
End Select

Exit Sub
Err_particular:
    Call ShowError("FilterOnDates")
    Resume Next
End Sub

Public Sub FilterEmpty()        '' Filters Records that are Empty
On Error GoTo Err_particular
    ConMain.Execute "Delete from tbldata where strF1 is null "
Exit Sub
Err_particular:
    Call ShowError("FilterEmpty")
End Sub

Public Sub StartProcessing(ByRef grd As MSFlexGrid, ByVal strSelEmp As String, _
    Optional frmDaily As Form)    '' Starts the Main Processing Loops
'Public Sub StartProcessing(ByRef grd As MSFlexGrid, SC1 As ScriptControl, ByVal strSelEmp As String)
On Error GoTo Err_particular
Dim strEmpArr() As String, intTmp As Integer
strSelEmp = Replace(strSelEmp, "'", "")
strSelEmp = Replace(strSelEmp, ")", "")
strSelEmp = Replace(strSelEmp, "(", "")
strEmpArr = Split(strSelEmp, ",")
grd.Visible = True
grd.Redraw = True
grd.Refresh

Do While typDT.dtFrom <= typDT.dtTo
    frmDaily.Refresh
    DoEvents: DoEvents: DoEvents: DoEvents
    typVar.strDtTrn = CStr(typDT.dtFrom)
    '' Changes as on 26-06-2003
    If typPerm.blnIO Then
    blnCurrDtTrnFound = FindTable(MonthName(Month(typDT.dtFrom), True) & _
    Right(CStr(Year(typDT.dtFrom)), 2) & "Trn")
    Else
    blnCurrDtTrnFound = FindTable(MonthName(Month(typDT.dtFrom), True) & _
    Right(CStr(Year(typDT.dtFrom)), 2) & "Trn")
    End If
    If typPerm.blnDI Then
    blnCurrDtTrnFound = FindTable(MonthName(Month(typDT.dtFrom), True) & _
    Right(CStr(Year(typDT.dtFrom)), 2) & "DI")
    End If
    ''
    grd.TextMatrix(1, 0) = typVar.strDtTrn             '' Reflect the Processing Date
    grd.Refresh
    For intTmp = 0 To UBound(strEmpArr)
        frmDaily.Refresh
        grd.Refresh
        Call SetMiscVars                    '' Re-initialize Miscellaneous Variablex
        typVar.strEmpCodeTrn = strEmpArr(intTmp) ''adrsDumb("Empcode")
        Call SettypTR                       '' Set the type Time Range
        Call SettypDH                       '' Set the type Daily Hours
        '' Set OT & CO Vars
        Call SettypOTVars
        Call SettypCOVars
        Call SetBreakHours                  '' Set the type Break hours
        '' Start Processing
        Call FillEmptype(typVar.strEmpCodeTrn)         '' Get Employee Details
            If typDT.dtFrom < typEmp.dtJoin Then    '' if Employee has not Joined yet
                GoTo Loop_Employee
            End If                              '' Check if the Leave Date is Less then Pro. Date

    
            If typEmp.dtLeft <= DateCompDate(typDT.dtFrom) Then
                GoTo Loop_Employee
            End If
        grd.TextMatrix(1, 1) = typVar.strEmpCodeTrn    '' Reflect the Processing Employee
        frmDaily.Refresh
        grd.Refresh
        Call GetLostPunches                 '' Get Lost Punches
        Call FilterOnTime                   '' Filter Punches on Time Basis
        '' Call FillCattype(typEmp.strECat)     '' Get Category Details
        Call FillOTType(typOTVars.bytOTCode)
        Call FillCOType(typCOVars.bytCOCode)
        Call FillCattype(typEmp.strECat)     '' Get Category Details
        If Day(DateCompDate(typDT.dtFrom)) = 1 Then
             Call FullCrLeave("Daily", typVar.strEmpCodeTrn, typEmp.strECat, typDT.dtFrom, typDT.dtFrom) 'added by
         End If
    
        Call GetPunchesArray                        '' Gets's the Array of Punches for two Days
        typVar.strLeaveStatus = GetLeaveStatus     '' Gest's the Leave Status
        If Not CheckShiftPunches() Then       '' Checks for Required Shifts and Punches
            GoTo Loop_Employee
        End If
        GetStatus (1)
        
        Select Case typEmp.bytEntry         '' Depending Upon the Entry Required
            Case 0      '' 0 Entry
                Call GetHoursZeroEnt
                GoTo Loop_Employee
            Case 1      '' 1 Entry
                Call GetLateHours ' MIS2007DF011 add by
                Call GetHoursOneEnt
                GoTo Loop_Employee
            Case Else
                '' Do Nothing
        End Select
        Select Case typVar.bytTmpEnt               '' Depending Upon the Entries Found
            Case 1, 3, 5, 7, 9              '' Odd Entries
                GetLateHours (1)          '' Get Late Hours

                GetPresent
                GetIrrMark
          
                Call AddRecordsToTrn
                GoTo Loop_Employee
            Case Else
                '' Do Nothing
        End Select
        Call ProcessHours
        Call PutHours                       '' Put Late,Early and COHrs in the Database
        Call AddRecordsToTrn
        
        '' End
Loop_Employee:
        '' adrsDumb.MoveNext                   '' Move to Next Employee
    Next

      typDT.dtFrom = typDT.dtFrom + 1

    frmDailyTry.Refresh
Loop
grd.Visible = False
grd.Refresh
Exit Sub
Err_particular:
    Call ShowError("StartProcessing")
End Sub
'Added by      'For monthly leave credit option for full and proposnate credit
Public Function FullCrLeave(ByVal ProcessType As String, ByVal STRECODE As String, ByVal strECat As String, ByVal Fdate As String, ByVal CreditDt As String, Optional ByVal Present As Integer)
On Error GoTo Err
Dim rsMChk As New ADODB.Recordset, rsEmp As New ADODB.Recordset

If rsMChk.State = 1 Then rsMChk.Close
rsMChk.Open "select lvcode,lv_qty,lv_acumul,crmonthly,fulcredit from leavdesc where cat='" & strECat & "'", ConMain
If rsEmp.State = 1 Then rsEmp.Close
rsEmp.Open "Select Joindate from Empmst where Empcode ='" & STRECODE & "'", ConMain

If Not (rsMChk.BOF And rsMChk.EOF) Then
    rsMChk.MoveFirst
    While Not (rsMChk.EOF)
        If (UCase(rsMChk.Fields("crmonthly")) = "Y") Then
            If ((UCase(rsMChk.Fields("fulcredit")) = "Y") And (Day(DateCompDate(Fdate)) = 1) And (ProcessType = "Daily")) Or _
            (UCase(rsMChk.Fields("fulcredit")) = "N" And ProcessType = "Monthly") Then
                strLeaveCode = rsMChk.Fields("lvcode"): sngLeaveAccu = rsMChk.Fields("lv_acumul")
                If ALreadyCreditDate(DateCompDate(CreditDt), STRECODE, strLeaveCode) Then
                    If DateAdd("M", -1, DateCompDate(Fdate)) < DateCompDate(rsEmp.Fields("joindate")) Then  'join newly
                        'check first 15 days AND first 15 days get half credit
                        If DateAdd("M", -1, DateCompDate(Fdate)) <= DateCompDate(rsEmp.Fields("joindate")) And _
                            DateCompDate(rsEmp.Fields("joindate")) <= DateAdd("D", -15, DateCompDate(Fdate)) Then
                                strTotalDays = CInt(CDbl(rsMChk.Fields("lv_qty")) / 2)
                        Else 'or else get only one fourth of total credit
                            strTotalDays = CInt((CDbl(rsMChk.Fields("lv_qty")) / 2) / 2)
                        End If
                    Else
                        strTotalDays = CDbl(rsMChk.Fields("lv_qty"))  'for old employee
                    End If
                    'check whether leave accum elizible or not
                    If sngLeaveAccu > GetLeaveBalance(STRECODE, strLeaveCode, CreditDt) Then
                        If (UCase(rsMChk.Fields("fulcredit")) = "Y") Then
                            Call LeaveCredit(STRECODE, CreditDt, strTotalDays, DateCompDate(FdtLdt(Month(CreditDt), Year(CreditDt), "l")))
                        ElseIf UCase(rsMChk.Fields("fulcredit")) = "N" Then
                            If Present >= 15 Then
                                Call LeaveCredit(STRECODE, CreditDt, strTotalDays, DateCompDate(FdtLdt(Month(CreditDt), Year(CreditDt), "f")))
                            Else
                                Call LeaveCredit(STRECODE, CreditDt, CInt(strTotalDays \ 2), DateCompDate(FdtLdt(Month(CreditDt), Year(CreditDt), "f")))
                            End If
                        End If
                    End If
                End If
            End If
        End If
        rsMChk.MoveNext
    Wend
End If
Exit Function
Err:
    ShowError ("FullCrLeave::Line=" & Erl)
    'Resume Next
End Function


Public Sub LeaveCredit(ByVal strEmpCode As String, ByVal strDate As String, ByVal strDay, Optional ByVal strTempdt As String)
        On Error GoTo Err
          Dim strquery As String, strRW As String
          Dim adrsTemp As New ADODB.Recordset, rsDel As New ADODB.Recordset
          
        strRW = "W"
        'Delete Earlier Credit From LvInfo and LvBal
        If rsDel.State = 1 Then rsDel.Close
        rsDel.Open "select * from lvinfo" & Right(pVStar.YearSel, 2) & " where fromdate=" & _
            strDTEnc & strTempdt & strDTEnc & " and todate=" & strDTEnc & strTempdt & strDTEnc & _
            " and empcode='" & strEmpCode & "' and trcd='2' and lcode='" & strLeaveCode & _
            "' and fromdate > " & strDTEnc & Date & strDTEnc, ConMain
        If Not (rsDel.BOF And rsDel.EOF) Then
            ConMain.Execute "update lvbal" & Right(pVStar.YearSel, 2) & _
                " set " & strLeaveCode & " =" & strLeaveCode & "-" & rsDel.Fields("days") & _
                " where empcode='" & strEmpCode & "'"
            ConMain.Execute "delete from lvinfo" & Right(pVStar.YearSel, 2) & _
                " where fromdate=" & strDTEnc & strTempdt & strDTEnc & " and todate=" & strDTEnc & _
                strTempdt & strDTEnc & " and empcode='" & strEmpCode & "' and trcd='2' and lcode='" & _
                strLeaveCode & "' and fromdate > " & strDTEnc & Date & strDTEnc
        End If
          '' Insert Information in LvInfo
        ConMain.Execute "INSERT INTO LvInfo" & Right(pVStar.YearSel, 2) & _
          "(Empcode,trcd,fromdate,todate,lcode,days,lv_type_rw,entrydate) values" & _
          "(" & "'" & strEmpCode & "'" & "," & "2" & "," & strDTEnc & DateSaveIns(strDate) & _
          strDTEnc & "," & strDTEnc & DateSaveIns(strDate) & strDTEnc & "," & "'" & _
          strLeaveCode & "'" & "," & strDay & "," & "'" & strRW & _
          "'" & "," & "'" & DateSaveIns(CStr(Date)) & "'" & ")"
          '' Update balance in LvBal
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT " & strLeaveCode & " FROM Lvbal" & _
          Right(pVStar.YearSel, 2) & " WHERE Empcode='" & strEmpCode & "'", ConMain
        If Not (adrsTemp.EOF And adrsTemp.BOF) Then
            If IsNull(adrsTemp.Fields("" & strLeaveCode & "")) Then
                ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
                  strLeaveCode & "=" & strDay & " where Empcode='" & strEmpCode & "'"
           Else
               ConMain.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & _
                  strLeaveCode & "=" & strLeaveCode & " +" & strDay & " where Empcode='" & strEmpCode & "'"
           End If
       End If
       Exit Sub
Err:
          ShowError ("LeaveCredit::Line=" & Erl)
End Sub


Private Function ALreadyCreditDate(strDate As String, _
    strEmpCode As String, strLeaveCode As String) As Boolean     '' Checks if Leave is Already Credited
On Error GoTo ERR_P                                 '' to the Employee for the Same Dates
Dim tmpDate As Date

ALreadyCreditDate = True
Dim strA_R As String
If Not FindTable("LvInfo" & Right(pVStar.YearSel, 2)) Then
    MsgBox "Leave file not present please create", vbInformation
    ALreadyCreditDate = False
    Exit Function
End If
tmpDate = Format(CDate(strDate), "DD/MMM/YYYY")
strA_R = "Select * from LvInfo" & Right(pVStar.YearSel, 2) & " where  FromDate=" & strDTEnc & _
" " & tmpDate & "" & strDTEnc & " and ToDate=" & strDTEnc & tmpDate & _
strDTEnc & " and Trcd= 2" & " and Empcode=" & "'" & strEmpCode & "'" & " and Lcode=" & _
"'" & strLeaveCode & "'"
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open strA_R, ConMain
If Not (adrsPaid.EOF And adrsPaid.BOF) Then
        ALreadyCreditDate = False
        Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ALreadyCreditDate :: ")
    ALreadyCreditDate = False
End Function


Public Function GetLeaveBalance(strEmpCode As String, strLeaveCode As String, _
    strDate As String) As Single
    Dim adrsTemp As Recordset
    Dim strLeaveTable As String
On Error GoTo GetLeaveBalance_Error
    strLeaveTable = "Lvbal" & Right(GetTrnYear(CDate(strDate)), 2)
    Set adrsTemp = OpenRecordSet("SELECT " & strLeaveCode & _
    " FROM " & strLeaveTable & " WHERE empcode='" & strEmpCode & "'")
    If Not (adrsTemp.EOF And adrsTemp.BOF) Then
        GetLeaveBalance = FilterNull(adrsTemp.Fields(strLeaveCode), NumericD)
    End If
On Error GoTo 0
Exit Function
GetLeaveBalance_Error:
   If Erl = 0 Then
      ShowError "Error in procedure GetLeaveBalance of Module mdlDaily"
   Else
      ShowError "Error in procedure GetLeaveBalance of Module mdlDaily And Line:" & Erl
   End If
End Function
Public Sub GetStatus(Optional bytR2 As Integer = 1) '' Calculates the Status of The Day
On Error GoTo Err_particular
Select Case bytR2
    Case 1              '' Get Basic Status
        typVar.strStatus = typVar.strLeaveStatus
        If Trim(typVar.strStatus) = "" Then
            Select Case typEmp.bytEntry
                Case 0
                    If typVar.strShiftTmp = typVar.strHlsCode Then
                        typVar.strStatus = typVar.strHlsCode & typVar.strHlsCode
                    ElseIf typVar.strShiftTmp = typVar.strWosCode Then
                        typVar.strStatus = typVar.strWosCode & typVar.strWosCode
                    Else
                        typVar.strStatus = typVar.strPrsCode & typVar.strPrsCode
                    End If

                Case Else
                    If typVar.strShiftTmp = typVar.strHlsCode Then
                        typVar.strStatus = typVar.strHlsCode & typVar.strHlsCode
                    ElseIf typVar.strShiftTmp = typVar.strWosCode Then
                        typVar.strStatus = typVar.strWosCode & typVar.strWosCode
                    Else
                        Select Case typVar.bytTmpEnt
                            Case 0
                                typVar.strStatus = typVar.strAbsCode & typVar.strAbsCode
                            Case Else
                                typVar.strStatus = typVar.strPrsCode & typVar.strPrsCode
                            End Select
                    End If
            End Select
        Else
            If typVar.strTmpLvtype = "W" Then
                If typVar.strShiftTmp = typVar.strHlsCode Then
                    typVar.strStatus = typVar.strHlsCode & typVar.strHlsCode
                End If
                If typVar.strShiftTmp = typVar.strWosCode Then
                    typVar.strStatus = typVar.strWosCode & typVar.strWosCode
                End If
            End If
        End If
        If InStr(typVar.strStatus, "  ") > 0 Then
            If typEmp.bytEntry = 0 Then
                typVar.strStatus = Replace(typVar.strStatus, "  ", typVar.strPrsCode)
            Else
                If typVar.bytTmpEnt = 0 Then
                    typVar.strStatus = Replace(typVar.strStatus, "  ", typVar.strAbsCode)
                Else
                    typVar.strStatus = Replace(typVar.strStatus, "  ", typVar.strPrsCode)
                End If
            End If
        End If
    Case 2          '' Get Status Based on his Late Hours OR Early Hours
        If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode Then
            If typDH.sngLateHrs >= typCat.sngCutL And typCat.sngCutL > 0 And _
            (typVar.strAflg <> "1" And typVar.strAflg <> "3") Then
                If IsNull(typVar.strStatus) Or Left(typVar.strStatus, 2) = typVar.strAbsCode Or _
                Left(typVar.strStatus, 2) = typVar.strPrsCode Then
                        typVar.strStatus = StuffVal(typVar.strStatus, 1, 2, typVar.strAbsCode)
                        typVar.blnTrnLate = True
                End If
            Else
                If IsNull(typVar.strStatus) Or Left(typVar.strStatus, 2) = typVar.strAbsCode Then
                    typVar.strStatus = StuffVal(typVar.strStatus, 1, 2, typVar.strPrsCode)
                End If
            End If
        Else
            If IsNull(typVar.strStatus) Or Left(typVar.strStatus, 2) = typVar.strAbsCode Then
                    typVar.strStatus = StuffVal(typVar.strStatus, 1, 2, typVar.strPrsCode)
            End If
        End If
        If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode Then
            If typDH.sngEarlyHrs >= typCat.sngCutE And typCat.sngCutE > 0 And typVar.strDflg <> "2" Then
                If IsNull(typVar.strStatus) Or Right(typVar.strStatus, 2) = typVar.strPrsCode Then
                    typVar.strStatus = StuffVal(typVar.strStatus, 3, 2, typVar.strAbsCode)
                    typVar.blnTrnErl = True
                End If
            Else
                If Right(typVar.strStatus, 2) = typVar.strAbsCode Or IsNull(Right(typVar.strStatus, 2)) Then
                    typVar.strStatus = StuffVal(typVar.strStatus, 3, 2, typVar.strPrsCode)
                End If
            End If
        Else
            If Right(typVar.strStatus, 2) = typVar.strAbsCode Or IsNull(Right(typVar.strStatus, 2)) Then
                typVar.strStatus = StuffVal(typVar.strStatus, 3, 2, typVar.strPrsCode)
            End If
        End If
    Case 3      '' When Not Hl /Wo Normal Case
        If IsNull(typVar.strStatus) Or Left(typVar.strStatus, 2) = typVar.strAbsCode Then
            typVar.strStatus = StuffVal(typVar.strStatus, 1, 2, typVar.strPrsCode)
        End If
    Case 4      '' When Not Hl /Wo Late Hour Checking
        If IsNull(typVar.strStatus) Or Left(typVar.strStatus, 2) = typVar.strAbsCode Then
            typVar.strStatus = StuffVal(typVar.strStatus, 1, 2, typVar.strPrsCode)
        End If
    Case 5      '' When Not Hl /Wo
        If IsNull(typVar.strStatus) Or Right(typVar.strStatus, 2) = typVar.strAbsCode Then
            typVar.strStatus = StuffVal(typVar.strStatus, 3, 2, typVar.strPrsCode)
        End If
    Case 6      '' When Not Hl /Wo
        If IsNull(typVar.strStatus) Or Right(typVar.strStatus, 2) = typVar.strAbsCode Then
            typVar.strStatus = StuffVal(typVar.strStatus, 3, 2, typVar.strPrsCode)
        End If
    Case 7      '' When not Late and Not Early
        If (typVar.blnTrnErl And typVar.blnTrnLate) Then typVar.strStatus = typVar.strAbsCode & typVar.strAbsCode
    Case 8      '' When not Late and not Early Early and Depending on Work Hours and 1/2 Day.
        If Not typVar.blnTrnLate And Not typVar.blnTrnErl And typDH.sngWorkHrs <= (typShift.sngHRS / 2) Then
            If UCase(typShift.strShift) = "O" Then Exit Sub
            If typVar.strStatus = ReplicateVal(typVar.strAbsCode, 2) Or typVar.strStatus = ReplicateVal(typVar.strPrsCode, 2) Then
                If typTR.sngTimeIn <= typShift.sngHalfE Then
                    typVar.strStatus = typVar.strPrsCode & typVar.strAbsCode
                Else
                    typVar.strStatus = typVar.strAbsCode & typVar.strPrsCode
                End If
            End If

        End If
    Case 9      '' When 1 Entry is required and Second Half is Found
        typVar.strStatus = typVar.strPrsCode & Right(typVar.strStatus, 2)
    Case 10     '' When 1 Entry is required and First Half is Found
        typVar.strStatus = Left(typVar.strStatus, 2) & typVar.strPrsCode
    Case 11
        typVar.strStatus = typVar.strAbsCode & typVar.strAbsCode
    Case 12 ''Minimum Half Day and Full Day Hours
        If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode Then
            If typDH.sngWorkHrs < AdrsCat.Fields("HalfDayHr") Then
                typVar.strStatus = typVar.strAbsCode & typVar.strAbsCode
            ElseIf typDH.sngWorkHrs < AdrsCat.Fields("FullDayHr") Then
                typVar.strStatus = typVar.strPrsCode & typVar.strAbsCode
            End If
        End If
End Select

Exit Sub
Err_particular:
    Call ShowError("GetStatus")
End Sub

Public Sub GetShiftOfDay(ByVal strDtShf As String, ByVal strEmp As String) '' Calculates the
On Error GoTo Err_particular                                               '' Shift of the
Dim strTmp As String, strnew As Date                                                  '' Day
Dim adrsDumb As New ADODB.Recordset
strTmp = Left(MonthName(Month(DateCompDate(strDtShf))), 3) & _
IIf(Month(DateCompDate(strDtShf)) < Val(pVStar.Yearstart), Right(pVStar.YearSel + 1, 2), _
Right(pVStar.YearSel, 2)) & "shf"
strnew = strDtShf
If adrsDumb.State = 1 Then adrsDumb.Close
If FindTable(strTmp) Then
    adrsDumb.Open "Select  D" & Day(DateCompDate(strDtShf)) & " from " & strTmp & _
    " where Empcode='" & strEmp & "'", ConMain
    If Not (adrsDumb.EOF And adrsDumb.BOF) Then
        typVar.strShiftOfDay = IIf(IsNull(adrsDumb(0)), "", adrsDumb(0))
    Else
        typVar.strShiftOfDay = ""
    End If
Else
    typVar.strShiftOfDay = ""
End If
typVar.strShiftTmp = typVar.strShiftOfDay
Exit Sub
Err_particular:
    Call ShowError("GetShiftOfDay")
    typVar.strShiftOfDay = ""
End Sub

Public Sub GetPresent()     ''Calculates the Present Display
If typVar.strStatus = "" Then
    typVar.sngPresent = 1
Else
    If Left(typVar.strStatus, 2) = Right(typVar.strStatus, 2) Then
        typVar.sngPresent = 1
    Else
        typVar.sngPresent = 0.5
    End If
End If

If GetFlagStatus("ActualPunch") Then
    If typTR.sngTimeOut <> 0 And typTR.sngTimeIn <> 0 Then
        typDH.sngWorkHrs = TimDiff(typTR.sngTimeOut, typTR.sngTimeIn)
    End If
End If

End Sub

Public Sub GetIrrMark()     '' Checks For Irregular Entries and Puts Irregular Marks
typVar.strIrrMark = ""
Select Case typEmp.bytEntry
    Case 0
    Case 1
    Case Else
        If typVar.bytTmpEnt > 0 And typVar.bytTmpEnt < typEmp.bytEntry Then typVar.strIrrMark = "*"
        If (typVar.bytTmpEnt > typEmp.bytEntry) And (typVar.bytTmpEnt Mod 2) <> 0 Then
            typVar.strIrrMark = "*"
        End If
        
        If blnIrregular And typPerm.blnIO Then typVar.strIrrMark = "*"
        ''
End Select
End Sub
Public Sub GetLunchLateHours(Optional bytR2 As Integer = 1) '' Calculates the Lunch Late Hours
Select Case bytR2
    Case 1
        If typBH.sngBrk1 <> 0 Then
            typBH.sngBrk1Late = TimDiff(typBH.sngBrk1, typShift.sngBH1)
            If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode Then
                If typBH.sngBrk1Late > 0 And typBH.sngBrk1Late <= typCat.sngLunchLtIgnore Then typBH.sngBrk1Late = 0
                If typBH.sngBrk1Late < 0 Then typBH.sngBrk1Late = 0
            Else
                typBH.sngBrk1Late = 0
            End If
        End If
End Select
Exit Sub
ERR_P:
    Call ShowError("GetLunchLateHours")
End Sub

Public Sub GetLateHours(Optional bytR2 As Integer = 1) '' Calculates the Late Hours
Dim tmpLate As Single
Dim arrLate() As String
On Error GoTo ERR_P
Select Case bytR2
    Case 1                          '' Get Basic Late Hours
            If typTR.sngTimeIn <> 0 Then
                typDH.sngLateHrs = TimDiff(typTR.sngTimeIn, typShift.sngIN)
              End If
            If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode Then
                If typDH.sngLateHrs > 0 And typDH.sngLateHrs <= typCat.sngLate Then
                         typDH.sngLateHrs = 0
                End If
                If typDH.sngLateHrs < 0 And Abs(typDH.sngLateHrs) <= typCat.sngEarlI Then _
                typDH.sngLateHrs = 0
            Else
                typDH.sngLateHrs = 0
            End If
    Case 2      '' Set Late Hours to Zero
        typDH.sngLateHrs = 0
    Case 3      '' If No Shift is Found
        If typTR.sngTimeIn > 0 Then typDH.sngLateHrs = typTR.sngTimeIn
    Case 4      '' When Single Entry is Required
        typDH.sngLateHrs = TimDiff(typTR.sngTimeIn, typShift.sngIN)
        
End Select
    If UCase(typShift.strShift) = "O" Then typDH.sngLateHrs = 0          '' For O Shift
Exit Sub
ERR_P:
    Call ShowError("GetLateHours")
End Sub

Public Sub GetEarlyHours(Optional bytR2 As Integer = 1)  '' Calculates the Early Hours
On Error GoTo ERR_P
Select Case bytR2
    Case 1      '' Basic Early Hours
            Select Case typVar.bytTmpEnt
                Case 1
                    typDH.sngEarlyHrs = 0
                Case Else
                    ''Case 2, 4, 6, 8
                    Select Case typShift.strShift       '' For O Shift
                        Case "O", "o"
                            typDH.sngEarlyHrs = 0 'TimDiff(typShift.sngHRS, TimDiff(typTR.sngTimeOut, typTR.sngTimeIn))
                        Case Else
                            If typTR.sngTimeOut <> 0 Then typDH.sngEarlyHrs = TimDiff(typShift.sngOut, _
                            typTR.sngTimeOut)
                    End Select
            End Select
            If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode Then
                If typDH.sngEarlyHrs > 0 And typDH.sngEarlyHrs <= typCat.sngEarl Then
                        typDH.sngEarlyHrs = 0
                End If
                If typDH.sngEarlyHrs <= 0 And Abs(typDH.sngEarlyHrs) <= typCat.sngLateI Then _
                typDH.sngEarlyHrs = 0
        
            Else
                typDH.sngEarlyHrs = 0
            End If
    Case 2      '' Set Early Hours to Zero
        typDH.sngEarlyHrs = 0
    Case 3      '' When No Shift
        If typTR.sngTimeOut > 0 Then typDH.sngEarlyHrs = typTR.sngTimeOut
End Select
Exit Sub
ERR_P:
    Call ShowError("GetEarlyHours")
End Sub

Public Sub GetBreak1()          '' Gets the First break
Select Case typVar.bytTmpEnt
   Case 3, 4, 5, 6, 7, 8
        If typTR.sngBreakS <> 0 And typTR.sngBreakE <> 0 Then
        typBH.sngBrk1 = TimDiff(typTR.sngBreakS, typTR.sngBreakE)
        End If
    End Select
End Sub

Public Sub GetBreak2()          '' Gets the Second Break
Select Case typVar.bytTmpEnt
    Case 6, 7, 8
        If typTR.sngTime5 <> 0 And typTR.sngTime6 <> 0 Then
        typBH.sngBrk2 = TimDiff(typTR.sngTime6, typTR.sngTime5)
        End If
End Select
End Sub

Public Sub GetBreak3()          '' Gets the Third Break
Select Case typVar.bytTmpEnt
    Case 8
        If typTR.sngTime7 <> 0 And typTR.sngTime8 <> 0 Then
        typBH.sngBrk3 = TimDiff(typTR.sngTime8, typTR.sngTime7)
        End If
End Select
End Sub

Public Sub GetTotalBreakHour()  '' Gets the Total BreakHours
    typTR.sngBreakHrs = TimAdd(typBH.sngBrk1, TimAdd(typBH.sngBrk2, typBH.sngBrk3))
End Sub

Public Sub GetOvertimeHours()  '' Calculates the Extra Work Hours
On Error GoTo ERR_P
Dim sngTmp As Single
'' If no rule then exit procedure
If typOTVars.bytOTCode = 100 Then Exit Sub
sngTmp = 0
'' Get Basic Hours based on the Day
Select Case typVar.strShiftTmp
    Case ""                     '' Weekday
        If typOTVars.bytOTWD = 1 Then
        sngTmp = TimDiff(typDH.sngWorkHrs, typShift.sngHRS)
        'sngTmp = TimDiff(TimDiff(typTR.sngTimeOut, typTR.sngTimeIn), typShift.sngHRS)
        Else
            Exit Sub
        End If
    Case typVar.strWosCode      '' Week Off
        If typOTVars.bytOTWO = 1 Then
            sngTmp = typDH.sngWorkHrs  ''TimDiff(typDH.sngWorkHrs, typShift.sngHRS)
        Else
            Exit Sub
        End If
    Case typVar.strHlsCode      '' Holiday
        If typOTVars.bytOTHL = 1 Then
            sngTmp = typDH.sngWorkHrs
        Else
            Exit Sub
        End If
    Case Else                   '' Weekday
        If typOTVars.bytOTWD = 1 Then
        sngTmp = TimDiff(typDH.sngWorkHrs, typShift.sngHRS)
        Else
            Exit Sub
        End If
End Select
'' If negative OT then exit sub
If sngTmp <= 0 Then Exit Sub
'' Adjust Late Early Hours
If typOTVars.bytDedLate = 0 Then
    '' Add back Late hours
    If typDH.sngLateHrs > 0 Then sngTmp = TimAdd(sngTmp, typDH.sngLateHrs)
End If
If typOTVars.bytDedEarl = 0 Then
    '' Add back Early hours
    If typDH.sngEarlyHrs > 0 Then sngTmp = TimAdd(sngTmp, typDH.sngEarlyHrs)
End If
'' Deductions
Select Case typVar.strShiftTmp
    Case ""                     '' Weekday
        sngTmp = DeductedOT(sngTmp)
    Case typVar.strWosCode      '' Week Off
        If typOTVars.bytApplyWO = 1 Then
            sngTmp = DeductedOT(sngTmp)
        End If
    Case typVar.strHlsCode      '' Holiday
        If typOTVars.bytApplyHL = 1 Then
            sngTmp = DeductedOT(sngTmp)
        End If
    Case Else                   '' Weekday
        sngTmp = DeductedOT(sngTmp)
End Select
'' If negative OT then exit sub
If sngTmp <= 0 Then Exit Sub
  If GetFlagStatus("DeductLunch") Or (GetFlagStatus("LESSBKHRSFROMOT") And typDH.sngWorkHrs < 12) Then  ' 10-02
    Dim LunchBrkHrs As Single
    LunchBrkHrs = 0
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select * from OTRul where otcode=" & typOTVars.bytOTCode, ConMain, adOpenStatic
                    
    Select Case typVar.strShiftTmp
        Case ""                     '' Weekday
           LunchBrkHrs = IIf(IsNull(adrsTemp.Fields("WDBkHr")), 0, adrsTemp.Fields("WDBkHr"))
        Case typVar.strWosCode      '' Week Off
           LunchBrkHrs = IIf(IsNull(adrsTemp.Fields("WOBkHr")), 0, adrsTemp.Fields("WOBkHr"))
        Case typVar.strHlsCode      '' Holiday
           LunchBrkHrs = IIf(IsNull(adrsTemp.Fields("HLBkHr")), 0, adrsTemp.Fields("HLBkHr"))
        Case Else                   '' Weekday
           LunchBrkHrs = IIf(IsNull(adrsTemp.Fields("WDBkHr")), 0, adrsTemp.Fields("WDBkHr"))
    End Select
    If sngTmp > LunchBrkHrs Then
        sngTmp = TimDiff(sngTmp, LunchBrkHrs)
    Else
        sngTmp = 0
    End If
End If

'' Round Off OT
sngTmp = RoundOffOT(sngTmp)

'' Check for Maximum OT
If sngTmp > typOTVars.sngMaxOT And typOTVars.sngMaxOT > 0 Then
    sngTmp = typOTVars.sngMaxOT
End If
typDH.sngOverTime = sngTmp
Exit Sub
ERR_P:
    Call ShowError("GetOvertimeHours")
End Sub

Private Function DeductedOT(ByVal sngTmp As Single) As Single
On Error Resume Next
Dim sngDedTmp As Single
sngDedTmp = sngTmp
'' Slab 1
If typOTVars.sngF1 > 0 Or typOTVars.sngT1 > 0 Then
    '' If OT Falls between the Slab
    If sngTmp >= typOTVars.sngF1 And sngTmp <= typOTVars.sngT1 Then
        '' If All OT is to be Deducted
        If typOTVars.bytAll1 Then
            sngDedTmp = 0
        Else
            sngDedTmp = TimDiff(sngTmp, typOTVars.sngD1)
        End If
    End If
End If
'' Slab 2
If typOTVars.sngF2 > 0 Or typOTVars.sngT2 > 0 Then
    '' If OT Falls between the Slab
    If sngTmp >= typOTVars.sngF2 And sngTmp <= typOTVars.sngT2 Then
        '' If All OT is to be Deducted
        If typOTVars.bytAll2 Then
            sngDedTmp = 0
        Else
            sngDedTmp = TimDiff(sngTmp, typOTVars.sngD2)
        End If
    End If
End If
'' Slab 3
If typOTVars.sngF3 > 0 Or typOTVars.sngT3 > 0 Then
    '' If OT Falls between the Slab
    If sngTmp >= typOTVars.sngF3 And sngTmp <= typOTVars.sngT3 Then
        '' If All OT is to be Deducted
        If typOTVars.bytAll3 Then
            sngDedTmp = 0
        Else
            sngDedTmp = TimDiff(sngTmp, typOTVars.sngD3)
        End If
    End If
End If
'' More Than
If typOTVars.sngMoreThan Then
    '' If OT is more than the Limit specified
    If sngTmp > typOTVars.sngMoreThan Then
        '' If All OT is to be Deducted
        If typOTVars.bytAll4 Then
            sngDedTmp = 0
        Else
            sngDedTmp = TimDiff(sngTmp, typOTVars.sngD4)
        End If
    End If
End If
'' Check for Negative Value
If sngDedTmp < 0 Then sngDedTmp = 0
'' Return the Deducted OT
DeductedOT = sngDedTmp
End Function

Private Function RoundOffOT(sngTmp As Single)
On Error Resume Next
Dim sngFrac As Single, sngTmp1 As Single
sngFrac = 0: sngTmp1 = 0
sngFrac = TimDiff(sngTmp, Fix(sngTmp))
If sngFrac > 0 Then
    sngTmp1 = sngFrac
    '' Between
    If sngFrac >= typOTVars.sngRF1 And sngFrac <= typOTVars.sngRT1 Then
        sngTmp1 = typOTVars.sngR1
    End If
    If sngFrac >= typOTVars.sngRF2 And sngFrac <= typOTVars.sngRT2 Then
        sngTmp1 = typOTVars.sngR2
    End If
    If sngFrac >= typOTVars.sngRF3 And sngFrac <= typOTVars.sngRT3 Then
        sngTmp1 = typOTVars.sngR3
    End If
    If sngFrac >= typOTVars.sngRF4 And sngFrac <= typOTVars.sngRT4 Then
        sngTmp1 = typOTVars.sngR4
    End If
    If sngFrac > typOTVars.sngRT5 And typOTVars.sngRT5 > 0 Then sngTmp1 = typOTVars.sngR5
    sngFrac = sngTmp1
    sngTmp = TimAdd(Fix(sngTmp), sngFrac)
End If
RoundOffOT = sngTmp
End Function

Public Sub PutLate()        '' Puts Late Hours in the LateErl Table
On Error GoTo ERR_P
Dim adrsDumb As New ADODB.Recordset
If typDH.sngLateHrs > 0 Then
    If (typDH.sngLateHrs < typCat.sngCutL Or typCat.sngCutL = 0) And _
    (typVar.strAflg <> "1" And typVar.strAflg <> "3") Then
        If adrsDumb.State = 1 Then adrsDumb.Close
        adrsDumb.Open "Select  * from LateErl where EmpCode=" & "'" & typEmp.strEmp & "'" & _
        " and " & strKDate & "=" & strDTEnc & DateCompStr(typDT.dtFrom) & strDTEnc, ConMain
        If Not (adrsDumb.EOF And adrsDumb.BOF) Then
            ConMain.Execute "Update LateErl set LateHrs=" & typDH.sngLateHrs & _
            " where Empcode=" & "'" & typEmp.strEmp & "'" & " and LateErl." & strKDate & "=" & strDTEnc & _
            DateCompStr(typDT.dtFrom) & strDTEnc
        Else
            ConMain.Execute "insert into LateErl (Empcode," & strKDate & ",Latehrs) values (" & _
            "'" & typEmp.strEmp & "'" & "," & strDTEnc & DateCompStr(typDT.dtFrom) & strDTEnc & _
            "," & typDH.sngLateHrs & ") "
        End If
    Else
        ConMain.Execute "Update LateErl Set LateHrs=0 Where Empcode='" & _
        typEmp.strEmp & "' and " & strKDate & "=" & strDTEnc & DateCompStr(typDT.dtFrom) & strDTEnc
    End If
End If
Exit Sub
ERR_P:
    Call ShowError("PutLate")
End Sub

Public Sub PutEarly()       '' Puts Early Hours in the LateErl Table
On Error GoTo Err_particular
Dim adrsDumb As New ADODB.Recordset
If typDH.sngEarlyHrs > 0 Then
    If (typDH.sngEarlyHrs < typCat.sngCutE Or typCat.sngCutE = 0) And typVar.strDflg <> "2" Then
        If adrsDumb.State = 1 Then adrsDumb.Close
        adrsDumb.Open "Select  * from LateErl where EmpCode=" & "'" & typEmp.strEmp & "'" & _
        " and " & strKDate & "=" & strDTEnc & DateCompStr(typDT.dtFrom) & strDTEnc, ConMain
        If Not (adrsDumb.EOF And adrsDumb.BOF) Then
            ConMain.Execute "Update LateErl set EarlHrs=" & typDH.sngEarlyHrs & _
            " where Empcode=" & "'" & typEmp.strEmp & "'" & " and LateErl." & strKDate & "=" & strDTEnc & _
            DateCompStr(typDT.dtFrom) & strDTEnc
        Else
            ConMain.Execute "insert into LateErl (Empcode," & strKDate & ",EarlHrs) values (" & _
            "'" & typEmp.strEmp & "'" & "," & strDTEnc & DateCompStr(typDT.dtFrom) & strDTEnc & _
            "," & typDH.sngEarlyHrs & ") "
        End If
    Else
        ConMain.Execute "Update LateErl Set EarlHrs=0 Where Empcode='" & _
        typEmp.strEmp & "' and " & strKDate & "=" & strDTEnc & DateCompStr(typDT.dtFrom) & strDTEnc
    End If
End If
Exit Sub
Err_particular:
    Call ShowError("PutEarly")
End Sub


Public Sub PutCOHrs()       '' Puts the COHrs in the Leave Information Table
On Error GoTo Err_particular
Dim adrsDumb As New ADODB.Recordset     '' adrsPaid
Dim adrsTmpCO As New ADODB.Recordset    '' adrsDept1
If typDH.sngCOHrs > 0 Then
    If FindTable("LvBal" & Right(pVStar.YearSel, 2)) Then
        If adrsDumb.State = 1 Then adrsDumb.Close
        adrsDumb.Open "Select LvCode,Run_Wrk From LeavDesc Where LvCode='CO' and Cat='" & _
        typCat.strCat & "'", ConMain, adOpenStatic
        If Not (adrsDumb.EOF And adrsDumb.BOF) Then
            If FieldExists("LvBal" & Right(pVStar.YearSel, 2), "CO") Then
                If adrsTmpCO.State = 1 Then adrsTmpCO.Close
                adrsTmpCO.Open "Select * from LvInfo" & Right(pVStar.YearSel, 2) & _
                " Where EmpCode='" & typEmp.strEmp & "' and FromDate=" & strDTEnc & _
                DateCompStr(typDT.dtFrom) & strDTEnc & " and LCode='" & _
                adrsDumb("LvCode") & "' and Trcd=2", ConMain
                If Not (adrsTmpCO.EOF And adrsTmpCO.BOF) Then
                    ConMain.Execute "Delete from LvInfo" & Right(pVStar.YearSel, 2) & _
                    " Where EmpCode='" & typEmp.strEmp & "' and FromDate=" & strDTEnc & _
                    DateCompStr(typDT.dtFrom) & strDTEnc & " and LCode='" & _
                    adrsDumb("LvCode") & "' and Trcd=2", ConMain
                    ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                    " Set CO=CO-" & IIf(IsNull(adrsTmpCO("days")), "0", adrsTmpCO("days")) & _
                    " Where EmpCode='" & typEmp.strEmp & "'" & " and CO is not NULL"
                    ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                    " Set CO=" & IIf(IsNull(adrsTmpCO("days")), "0", adrsTmpCO("days")) & _
                    " Where EmpCode='" & typEmp.strEmp & "'" & " and CO is NULL"
                End If
                ConMain.Execute "insert into LvInfo" & _
                Right(pVStar.YearSel, 2) & "(Empcode,Trcd,FromDate,ToDate,LCode," & _
                "Days,Lv_type_RW,EntryDate) Values('" & typEmp.strEmp & "',2," & _
                strDTEnc & DateSaveIns(typDT.dtFrom) & strDTEnc & "," & strDTEnc & _
                DateSaveIns(typDT.dtFrom) & strDTEnc & ",'CO'," & typDH.sngCOHrs & ",'" & _
                adrsDumb("Run_Wrk") & "'," & strDTEnc & DateSaveIns(typDT.dtFrom) & _
                strDTEnc & ")"
                ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                " Set CO=CO+" & typDH.sngCOHrs & " Where EmpCode='" & typEmp.strEmp & "'" & _
                " and CO is not NULL"
                ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                " Set CO=" & typDH.sngCOHrs & " Where EmpCode='" & typEmp.strEmp & "'" & _
                " and CO is NULL"
                
            Else
                MsgBox NewCaptionTxt("M3001", adrsMod), vbExclamation
 
            End If
        Else
                MsgBox NewCaptionTxt("M3002", adrsMod), vbExclamation
        End If
    Else
            MsgBox NewCaptionTxt("M3003", adrsMod) & vbCrLf & _
            NewCaptionTxt("M3004", adrsMod), vbExclamation
            
    End If
Else
'' 24/12/04
'' ZF Pune
''
'' If rec is avail. then delete
   If FindTable("LvBal" & Right(pVStar.YearSel, 2)) Then
        If adrsDumb.State = 1 Then adrsDumb.Close
        adrsDumb.Open "Select LvCode,Run_Wrk From LeavDesc Where LvCode='CO' and Cat='" & _
        typCat.strCat & "'", ConMain, adOpenStatic
        If Not (adrsDumb.EOF And adrsDumb.BOF) Then
            If FieldExists("LvBal" & Right(pVStar.YearSel, 2), "CO") Then
                If adrsTmpCO.State = 1 Then adrsTmpCO.Close
                adrsTmpCO.Open "Select * from LvInfo" & Right(pVStar.YearSel, 2) & _
                " Where EmpCode='" & typEmp.strEmp & "' and FromDate=" & strDTEnc & _
                DateCompStr(typDT.dtFrom) & strDTEnc & " and LCode='" & _
                adrsDumb("LvCode") & "' and Trcd=2", ConMain
                If Not (adrsTmpCO.EOF And adrsTmpCO.BOF) Then
                    ConMain.Execute "Delete from LvInfo" & Right(pVStar.YearSel, 2) & _
                    " Where EmpCode='" & typEmp.strEmp & "' and FromDate=" & strDTEnc & _
                    DateCompStr(typDT.dtFrom) & strDTEnc & " and LCode='" & _
                    adrsDumb("LvCode") & "' and Trcd=2", ConMain
                    ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                    " Set CO=CO-" & IIf(IsNull(adrsTmpCO("days")), "0", adrsTmpCO("days")) & _
                    " Where EmpCode='" & typEmp.strEmp & "'" & " and CO is not NULL"
                    ConMain.Execute "update LvBal" & Right(pVStar.YearSel, 2) & _
                    " Set CO=" & IIf(IsNull(adrsTmpCO("days")), "0", adrsTmpCO("days")) & _
                    " Where EmpCode='" & typEmp.strEmp & "'" & " and CO is NULL"
                End If
                End If
       End If
    End If
End If
Exit Sub
Err_particular:
    Call ShowError("PutCOHrs")
End Sub

Public Sub GetCOHrs()     '' Calculates the CO of the Day
On Error GoTo Err_particular
Dim sngTmp As Single, sngWrkTmp As Single
'' If no rule then exit procedure
If typCOVars.bytCOCode = 100 Then Exit Sub
sngTmp = 0: sngWrkTmp = 0
'' Get Basic Hours based on the Day
Select Case typVar.strShiftTmp
    Case ""                     '' Weekday
        If typCOVars.bytCOWD = 1 Then
            sngWrkTmp = TimDiff(typDH.sngWorkHrs, typShift.sngHRS)
        Else
            Exit Sub
        End If
    Case typVar.strWosCode      '' Week Off
        If typCOVars.bytCOWO = 1 Then
            'sngWrkTmp = TimDiff(typDH.sngWorkHrs, typShift.sngHRS)
            sngWrkTmp = typDH.sngWorkHrs '-- Supriya FairField 17/02/05 to remove standard error
        Else
            Exit Sub
        End If
    Case typVar.strHlsCode      '' Holiday
        If typCOVars.bytCOHL = 1 Then
            sngWrkTmp = typDH.sngWorkHrs
        Else
            Exit Sub
        End If
    Case Else                   '' Weekday
        If typCOVars.bytCOWD = 1 Then
            sngWrkTmp = TimDiff(typDH.sngWorkHrs, typShift.sngHRS)
        Else
            Exit Sub
        End If
End Select
'' If negative OT then exit sub
If sngWrkTmp <= 0 Then Exit Sub
'' Adjust Late Early Hours
If typCOVars.bytCOLate = 0 Then
    '' Add back Late hours
    If typDH.sngLateHrs > 0 Then sngWrkTmp = TimAdd(sngWrkTmp, typDH.sngLateHrs)
End If
If typCOVars.bytCOEarl = 0 Then
    '' Add back Early hours
    If typDH.sngEarlyHrs > 0 Then sngWrkTmp = TimAdd(sngWrkTmp, typDH.sngEarlyHrs)
End If
'' Calculate Compoff Day
Select Case typVar.strShiftTmp
    Case ""                         '' Weekday
        If sngWrkTmp >= typCOVars.sngWDH And typCOVars.sngWDH > 0 Then sngTmp = 0.5
        If sngWrkTmp >= typCOVars.sngWDF And typCOVars.sngWDF > 0 Then sngTmp = 1
        If typCOVars.bytCOWD = 0 Then sngTmp = 0
    Case typVar.strWosCode          '' Weekoff
        If sngWrkTmp >= typCOVars.sngWOH And typCOVars.sngWOH > 0 Then sngTmp = 0.5
        If sngWrkTmp >= typCOVars.sngWOF And typCOVars.sngWOF > 0 Then sngTmp = 1
        If typCOVars.bytCOWO = 0 Then sngTmp = 0
    Case typVar.strHlsCode          '' Holiday
        If sngWrkTmp >= typCOVars.sngHLH And typCOVars.sngHLH > 0 Then sngTmp = 0.5
        If sngWrkTmp >= typCOVars.sngHLF And typCOVars.sngHLF > 0 Then sngTmp = 1
        If typCOVars.bytCOHL = 0 Then sngTmp = 0
    Case Else                       '' Weekday
        If sngWrkTmp >= typCOVars.sngWDH And typCOVars.sngWDH > 0 Then sngTmp = 0.5
        If sngWrkTmp >= typCOVars.sngWDF And typCOVars.sngWDF > 0 Then sngTmp = 1
        If typCOVars.bytCOWD = 0 Then sngTmp = 0
End Select
typDH.sngCOHrs = sngTmp

Exit Sub
Err_particular:
    Call ShowError("GetCOHrs")
End Sub

Public Sub GetRemarks()     '' Makes Remarks based on Late and Early Hours
On Error GoTo Err_particular
'' No Remarks
typVar.strRemarks = ""
'' Early With Permission and Late with Permission
If (typDH.sngEarlyHrs > 0 And typVar.strDflg = "2") And (typDH.sngLateHrs > 0 And _
(typVar.strAflg = "1" Or typVar.strAflg = "3")) Then
    typVar.strRemarks = "E P/ L P"
End If
'' Early With Permission and Late Without Permission
If (typDH.sngEarlyHrs > 0 And typVar.strDflg = "2") And (typDH.sngLateHrs > 0 And _
(typVar.strAflg <> "1" And typVar.strAflg <> "3")) Then
    typVar.strRemarks = "E P/ L"
End If
'' Early With Permisson and not Late
If (typDH.sngEarlyHrs > 0 And typVar.strDflg = "2") And (typDH.sngLateHrs <= 0) Then
    typVar.strRemarks = "E P"
End If
'' Early Without Permission and Late With Permission
If (typDH.sngEarlyHrs > 0 And typVar.strDflg <> "2") And (typDH.sngLateHrs > 0 And _
(typVar.strAflg = "1" Or typVar.strAflg = "3")) Then
    typVar.strRemarks = "E/ L P"
End If
'' Early Without Permisson and Late Without Permission
If (typDH.sngEarlyHrs > 0 And typVar.strDflg <> "2") And (typDH.sngLateHrs > 0 And _
(typVar.strAflg <> "1" And typVar.strAflg <> "3")) Then
    typVar.strRemarks = "E/ L"
End If
'' Early Without Permission and not Late
If (typDH.sngEarlyHrs > 0 And typVar.strDflg <> "2") And (typDH.sngLateHrs <= 0) Then
    typVar.strRemarks = "E"
End If
'' Not Early and Late With Permission
If (typDH.sngEarlyHrs <= 0) And (typDH.sngLateHrs > 0 And _
(typVar.strAflg = "1" Or typVar.strAflg = "3")) Then
    typVar.strRemarks = "L P"
End If
'' Not Early and Late Without Permission
If (typDH.sngEarlyHrs <= 0) And (typDH.sngLateHrs > 0 And _
(typVar.strAflg <> "1" And typVar.strAflg <> "3")) Then
    typVar.strRemarks = "L"
End If
Exit Sub
Err_particular:
    Call ShowError("GetRemarks")
End Sub

Public Sub GetLostPunches() '' Gets the Punches from the Lost Entry Table
On Error GoTo Err_particular
Select Case bytBackEnd
Case 1, 2 ''Access ,SQL-SERVER
    
        If blnSP = False Then
            If bytBackEnd = 2 Then
                ConMain.Execute "insert into dailypro(dte,t_punch,Empcode,flg,card,SHIFT) select " & _
                "l." & strKDate & ", format(l.t_punch,'0.00') as t ,l.Empcode,'9',e.card,l.shift from lost l,empmst e where " & _
                " (l." & strKDate & " between " & strDTEnc & Format(typDT.dtFrom, "dd/mmm/yyyy") & strDTEnc & " and " & strDTEnc & _
                Format(typDT.dtTo + 1, "dd/mmm/yyyy") & strDTEnc & ") and l.Empcode='" & _
                typEmp.strEmp & "' and l.Empcode=e.Empcode"
                
            Else
                ConMain.Execute "insert into dailypro(Empcode, t_punch,dte,flg, card , shift)  select e.Empcode,  l.[Date]+convert(datetime, replace(format(l.t_punch, '00.00'),'.',':')), l.[Date], '9', E.card, L.shift from lost as l,empmst e where  (l.[Date] between ' " & Format(typDT.dtFrom, "dd/mmm/yyyy") & "' and '" & Format(typDT.dtTo + 1, "dd/mmm/yyyy") & "') and l.Empcode='" & typEmp.strEmp & "' and l.Empcode=e.Empcode"
            End If
                
                
'            Else
'
'            End If
        End If
Case 3 ''Oracle
    'shift column add by  MIS2007DF012
    ConMain.Execute "insert into dailypro(unqfld,dte,t_punch,Empcode,flg,card,SHIFT) select " & _
    "Next1.NextVal,l." & strKDate & ",l.t_punch,l.Empcode,'9',e.card,l.shift from lost l,empmst e where " & _
    " (l." & strKDate & " between " & strDTEnc & DateCompStr(typDT.dtFrom) & strDTEnc & " and " & strDTEnc & _
    DateCompStr(typDT.dtTo + 1) & strDTEnc & ") and l.Empcode='" & _
    typEmp.strEmp & "' and l.Empcode=e.Empcode"
    ''
End Select
Exit Sub
Err_particular:
    Call ShowError("GetLostPunches")
End Sub

Public Sub GetDataPunches(ByVal strSelEmp As String)         '' Gets the Punches into the Processing Table from the
On Error GoTo Err_particular    '' Raw Data Table
    If typPerm.blnIO = False Then
        If bytBackEnd = 2 Then
            strSql = "insert into dailypro(t_punch,dte,card) " & _
            "select   format(strf1,'hh.mm') as tPunch," & _
            " cdate(Format(strf1, 'dd/mmm/yyyy'))  as dt, strcode from tbldata"
        Else
'            strSql = "insert into dailypro(t_punch,dte,card) select   format(strf1,'hh.mm') as tPunch, Format(strf1, 'dd/mmm/yyyy')  as dt, strcode from tbldata"
          
                strSql = "insert into dailypro(t_punch,dte,card) select   strf1 as tPunch, strf1  as dt, strcode from tbldata"
                
        End If
        
     Else
        If bytBackEnd = 2 Then
            strSql = "insert into dailypro(t_punch,dte,card, Shift) " & _
            "select   format(strf1,'hh.mm') as tPunch," & _
            " cdate(Format(strf1, 'dd/mmm/yyyy'))  as dt, LEFT(strcode, " & pVStar.CardSize & ") , Right(strcode,1) as S from tbldata"
        Else
                strSql = "insert into dailypro(t_punch,dte,card, Shift) select   format(strf1,'hh.mm') as tPunch, Format(strf1, 'dd/mmm/yyyy')  as dt, LEFT(strcode, 4) , Right(strcode,1) as S from tbldata"
        End If
     End If
        
        ConMain.Execute strSql, i
        If bytBackEnd = 1 Then ConMain.Execute "UPDATE DailyPro SET Dte = cast([dte] as date) WHERE cast([dte] as time) > '00:00'", i
Select Case bytBackEnd
    Case 1 ''SQL Server
        ConMain.Execute "UPDATE Dailypro SET DailyPro.Empcode = " & _
        "Empmst.Empcode from dailypro INNER JOIN Empmst ON Empmst.Card = DailyPro.Card  " & _
        "Where Empmst.Empcode in " & strSelEmp
    Case 2 ''MS-ACCESS
        ConMain.Execute "UPDATE Dailypro INNER JOIN Empmst " & _
        "ON val(Empmst.Card) = val(DailyPro.Card) SET " & _
        "DailyPro.Empcode = Empmst.Empcode Where Empmst.Empcode in " & strSelEmp
    Case 3 '' ORACLE
        ConMain.Execute "UPDATE Dailypro a SET Empcode = (select " & _
        "Empmst.Empcode  from Empmst where empmst.Card = a.Card and Empmst.Empcode in " & _
        strSelEmp & ")"
End Select
Call FillMissedCards
''

ConMain.Execute "Delete from DailyPro Where " & _
"(Empcode is null or Empcode='') and Card not in ('" & typPerm.strOD & _
"','" & typPerm.strLate & "','" & typPerm.strEarl & "','" & _
typPerm.strBus & "')"


'case 3 ''Oracle
'End Select
'' Start Update for Next Years Data
'Dim dttmp As Date
'dttmp = typDT.dtTo + 1
'dttmp = DateAdd("YYYY", -1, dttmp)
'Select Case bytBackEnd
'    Case 1      '' SQL Server
'        conmain.Execute "Update DailyPro Set Dte=DateAdd(year,1,Dte) " & _
'        "where Dte<=" & strDTEnc & DateCompStr(dttmp) & strDTEnc
'    Case 2      '' MS-Access
'        conmain.Execute "Update DailyPro Set Dte=DateAdd('" & _
'        "YYYY',1,Dte) where Dte<=" & strDTEnc & DateCompStr(dttmp) & strDTEnc
'    Case 3      ''Oracle
'        conmain.Execute "Update DailyPro Set Dte=ADD_MONTHS(Dte,12) " & _
'        "where Dte<=" & strDTEnc & DateCompStr(dttmp) & strDTEnc
'End Select
'' End Update for Next Years Data
Exit Sub
Err_particular:
    Call ShowError("GetDataPunches")
'    Resume Next
End Sub

Public Sub FilterOnTime()   '' Filters the Data Based on Time
On Error GoTo Err_particular
Dim sngTmp As Single, sngT_Punch As Single, strDtDate As Date, strFlgTmp As String
Dim adrsDumb As New ADODB.Recordset, strDelStr As String
Dim sngArr() As Single
Dim NightDt As Date
sngTmp = -1
'' Delete Permission Records
    ConMain.Execute "Delete from dailypro where Empcode is Null"
'' End
If adrsDumb.State = 1 Then adrsDumb.Close
adrsDumb.Open "Select unqfld,dte,t_punch,Flg,Shift from DailyPro where Empcode='" & typEmp.strEmp & _
"' order by Empcode,dte,t_punch, Shift ", ConMain
If adrsDumb.EOF And adrsDumb.BOF Then Exit Sub

Dim rngIn As Date
Dim rngOut As Date
Dim rngDiff As Date
Dim HoldSeconds As Long
Dim InOut As String

Do While Not adrsDumb.EOF
Start_Loop:
    strDtDate = adrsDumb("dte")
    sngT_Punch = Format(adrsDumb("t_punch"), "HH.nn")
    strFlgTmp = adrsDumb("Flg")
    InOut = IIf(IsNull(adrsDumb("Shift")), "", adrsDumb("Shift"))
    adrsDumb.MoveNext
    If Not adrsDumb.EOF Then
        If adrsDumb("dte") = strDtDate And strFlgTmp = adrsDumb("Flg") Then
            If InOut = "O" Then
                If adrsDumb("SHIFT") = "I" Then
'                        HoldSeconds = DateDiff("s", sngT_Punch, adrsDumb.Fields("t_punch").Value)
'                        rngDiff = Format(DateAdd("s", HoldSecondadrsDumb.Fields("t_punch").Values, "00:00:00"), "hh:mm:ss")
'                        rngOut = DateAdd("s", HoldSeconds, rngOut)
'                        typTR.sngBreakE = DateAdd("m", typTR.sngBreakE, DateDiff("m", rngOut, sngT_Punch))
                        
                        typTR.sngBreakE = TimDiff(Format(adrsDumb.Fields("t_punch").Value, "HH.nn"), sngT_Punch)
                        typTR.sngBreakS = typTR.sngBreakE + typTR.sngBreakS
                End If
        
            End If
        
        
        End If
               
    
    
    
        If adrsDumb("dte") = strDtDate And strFlgTmp = adrsDumb("Flg") And _
        Val(Format(adrsDumb("t_punch"), "HH.nn")) < TimAdd(sngT_Punch, typPerm.sngFiltTime + 0.01) Then
            sngTmp = sngTmp + 1
            ReDim Preserve sngArr(sngTmp)
            sngArr(sngTmp) = adrsDumb("unqfld")
        Else
            GoTo Start_Loop
        End If
        
        
        
        
        

    End If
Loop
If sngTmp = -1 Then Exit Sub
For sngTmp = 0 To UBound(sngArr)
    ''conmain.Execute "Delete from Dailypro where unqfld=" & sngArr(sngTmp)
    strDelStr = strDelStr & sngArr(sngTmp) & ","
Next
strDelStr = Left(strDelStr, Len(strDelStr) - 1)
strDelStr = "(" & strDelStr & ")"
ConMain.Execute "Delete from Dailypro where unqfld in " & strDelStr
Erase sngArr
Exit Sub
Err_particular:
    Call ShowError("FilterOnTime")
    'Resume Next
End Sub


Public Sub PutFlag()        '' Procedure to set The Flags Based on Cards
On Error GoTo Err_particular
Dim strFlgTmp As String
Dim blnTmpBus As Boolean, blnTmpTemp As Boolean
Dim adrsDumb As New ADODB.Recordset
blnTmpBus = False: blnTmpTemp = False
If adrsDumb.State = 1 Then adrsDumb.Close
adrsDumb.Open "Select UnqFld,Card,Flg,EmpCode from Dailypro", _
ConMain, adOpenKeyset, adLockOptimistic
If (adrsDumb.EOF And adrsDumb.BOF) Then Exit Sub
If typPerm.blnIsPerm Then
    Do
        Select Case adrsDumb!card
            Case Is = typPerm.strOD
                strFlgTmp = "0"     '' Official Duty
                blnTmpBus = False
                Call SetRecordFlag(adrsDumb, strFlgTmp)
                adrsDumb.MoveNext
                Call SetRecordFlag(adrsDumb, strFlgTmp)
            Case Is = typPerm.strLate
                blnTmpBus = False
                strFlgTmp = "1"     '' Late
                Call SetRecordFlag(adrsDumb, strFlgTmp)
                adrsDumb.MoveNext
                Call SetRecordFlag(adrsDumb, strFlgTmp)
            Case Is = typPerm.strEarl
                blnTmpBus = False
                strFlgTmp = "2"     '' Early
                Call SetRecordFlag(adrsDumb, strFlgTmp)
                adrsDumb.MoveNext
                Call SetRecordFlag(adrsDumb, strFlgTmp)
            Case Is = typPerm.strBus
                strFlgTmp = "3"     '' Bus
                blnTmpBus = True
                Do While blnTmpBus And Not adrsDumb.EOF
                    Call SetRecordFlag(adrsDumb, strFlgTmp)
                    adrsDumb.MoveNext
                    
                    If adrsDumb.EOF Then
                        NoEndBusCard
                        Exit Do
                    Else
                        Select Case adrsDumb("card")
                            Case typPerm.strOD, typPerm.strLate, typPerm.strEarl
                                NoEndBusCard
                                Exit Do
                        End Select
                    End If
                    ''
                    If adrsDumb("card") = typPerm.strBus Then
                        Call SetRecordFlag(adrsDumb, strFlgTmp)
                        blnTmpBus = False
                    End If
                Loop
                ''blnTmpBus = True
            Case Else               '' Normal
                blnTmpBus = False
                strFlgTmp = "9"
                Call SetRecordFlag(adrsDumb, strFlgTmp)
                ''adrsDumb.MoveNext
        End Select
        If blnTmpBus Then
            ConMain.Execute "Update Dailypro Set Flg=" & "'" & 9 & "'" & _
            " where Flg= " & "'" & 3 & "'"
        End If
        If blnTmpTemp Then
            ConMain.Execute "Update Dailypro Set Flg=" & "'" & 9 & "'" & _
            " where Flg= " & "'" & 4 & "'"
        End If
        If blnTmpBus = False And blnTmpTemp = False Then
            If Not adrsDumb.EOF Then adrsDumb.MoveNext
        End If
    Loop Until adrsDumb.EOF
    adrsDumb.Close
Else     '' i.e. Permission Card = False
    strFlgTmp = "9"
    If adrsDumb.State = 1 Then adrsDumb.Close
    ConMain.Execute "Update DailyPro Set Flg='9'"
End If
Exit Sub
Err_particular:
    Call ShowError("PutFlag")
    ''Resume Next
End Sub

Private Sub NoEndBusCard()
    MsgBox "Bus end card not found", vbInformation
    ConMain.Execute "update Dailypro set Flg='9' where flg='3' and card <> '" & typPerm.strBus & "'"
End Sub

Public Sub SetRecordFlag(ByRef adrsRef As ADODB.Recordset, ByVal strFlagTmp As String)
On Error GoTo Err_particular    '' Function to Set the Flags of Individual Records
If adrsRef.EOF Then Exit Sub    '' Put in the Processing Table
If strFlagTmp = "9" Then
    adrsRef("flg") = strFlagTmp
Else
    If strFlagTmp = "3" Then
        If Not IsNull(adrsRef("EmpCode")) Then
            adrsEmp.MoveFirst
            adrsEmp.Find "Empcode='" & adrsRef("Empcode") & "'"
            
            If Not adrsEmp.EOF Then
                typEmp.strConv = IIf(IsNull(adrsEmp("Conv")), "", Trim(adrsEmp("Conv")))
            Else
                typEmp.strConv = ""
            End If
        Else
            typEmp.strConv = ""
            '' Call FillEmptype(adrsRef("Empcode"))
        End If
        If typEmp.strConv = "B" Or typEmp.strConv = "" Then
            adrsRef("flg") = strFlagTmp
        Else
            adrsRef("flg") = "9"
        End If
    ElseIf strFlagTmp = "4" Then
        adrsRef("flg") = strFlagTmp
    Else
        adrsRef("flg") = strFlagTmp
    End If
End If
adrsRef.Update
Exit Sub
Err_particular:
    Call ShowError("SetRecordFlag")

End Sub

Public Sub OpenMasters(Optional ByVal strSelEmp As String = "")  '' Opens All the Necessary Master Tables
On Error GoTo Err_particular
'' Open Category Master
If AdrsCat.State = 1 Then AdrsCat.Close

AdrsCat.Open "Select * from CatDesc where cat <> '100'", ConMain, adOpenStatic, adLockReadOnly

'' Open Employee Master
If strSelEmp <> "" Then
    strSelEmp = " where Empcode in " & strSelEmp
End If
If adrsEmp.State = 1 Then adrsEmp.Close


adrsEmp.Open "Select Empcode,Name,Shf_Chg,Entry,Cat,Card," & strKOff & ",Off2,Wo_1_3," & _
"Wo_2_4,Styp,Joindate,Leavdate,SCode,F_Shf,OTCode,COCode,WOHLAction,Action3Shift," & _
"AutoForPunch,ActionBlank,AutoG from Empmst" & strSelEmp _
, ConMain, adOpenStatic, adLockReadOnly

''
'' Open Shift Master
If adRsintshft.State = 1 Then adRsintshft.Close

adRsintshft.Open "Select * from instshft where shift <> '100'", ConMain, adOpenStatic, adLockReadOnly
'' Unuseable Recordsets
'' Open OT Recordset
If adrsOT.State = 1 Then adrsOT.Close
adrsOT.Open "Select * from OTRul", ConMain, adOpenStatic, adLockReadOnly
'' Open CO Recordsets
If adrsCO.State = 1 Then adrsCO.Close
adrsCO.Open "Select * from CORul", ConMain, adOpenStatic, adLockReadOnly
'' adrsCat, adrsEmp, adRsintshft
Exit Sub
Err_particular:
    Call ShowError("OpenMasters")
End Sub


Public Sub GetAutoShift(ByVal sngShiftTime As Single)    '' Gets the Auto Shift for a
On Error GoTo Err_particular                     '' Particular Employee on a Particular Day
Dim strTmpShift As String, bytTmp As Integer
Dim adrsDumb As New ADODB.Recordset
Dim sngOne As Single, sngTwo As Single
'' Get Shift into the Temporary Variable
strTmpShift = typVar.strShiftOfDay
'' Check if Late or Early Within the Time of Shift in Time
Call FillShifttype(typVar.strShiftOfDay)
If sngShiftTime > typShift.sngIN + typPerm.sngPostLt Then '' Late
    ''typVar.strShiftOfDay = RetNextAutoShift(sngShiftTime, typVar.strShiftOfDay)
    bytTmp = 1
ElseIf sngShiftTime < typShift.sngIN + typPerm.sngPostEarl Then  ''Early
    ''typVar.strShiftOfDay = RetPrevAutoShift(sngShiftTime, typVar.strShiftOfDay)
    bytTmp = 1
End If

If typVar.strShiftOfDay = "" Then bytTmp = 1
Dim strAuTmp() As String, bytCnt As Integer
Dim strT1 As String
 
        strT1 = ""
        If adrsDumb.State = 1 Then adrsDumb.Close
         If Left(UCase(typEmp.strAutoGroup), 3) = "ALL" Then
            adrsDumb.Open "Select Shift From instshft ", ConMain, adOpenDynamic, adLockOptimistic
            For i = 1 To adrsDumb.RecordCount
                strT1 = strT1 & "'" & LTrim(adrsDumb.Fields(0)) & "',"
                adrsDumb.MoveNext
            Next
        Else
            strT1 = Replace(typEmp.strAutoGroup, ".", "','")
            If strT1 <> "" Then
            strT1 = Left(strT1, Len(strT1) - 1)
            strT1 = "'" & strT1
            Else
            strT1 = "'" & strT1
            End If
        End If
        
        If Len(strT1) > 0 Then strT1 = Left(strT1, Len(strT1) - 1)
        If strT1 <> "" Then                                                     'modified by
        strT1 = " SHIFT IN (" & strT1 & ") AND "
        Else
        strT1 = " SHIFT IN ('" & strT1 & "') AND "
        End If
If bytTmp = 1 Then
    If adrsDumb.State = 1 Then adrsDumb.Close
    
    adrsDumb.Open "SELECT SHIFT,SHF_IN FROM INSTSHFT WHERE " & strT1 & " (SHF_IN IN (SELECT MIN(SHF_IN) " & _
    "AS A1 FROM INSTSHFT WHERE " & strT1 & " SHF_IN >=" & sngShiftTime & ") OR SHF_IN IN (SELECT  " & _
    "MAX(SHF_IN)  AS A1 FROM INSTSHFT WHERE " & strT1 & " SHF_IN <=" & sngShiftTime & ")) order by shift", ConMain, adOpenStatic, adLockReadOnly
     
     tempshift = typVar.strShiftOfDay
    If Not (adrsDumb.EOF And adrsDumb.BOF) Then
        If adrsDumb.RecordCount = 1 Then
            typVar.strShiftOfDay = adrsDumb("shift")
        Else
            sngOne = adrsDumb("shf_in")
            adrsDumb.MoveNext
            sngTwo = adrsDumb("shf_in")
            If sngTwo > sngShiftTime Then
                If TimDiff(sngTwo, sngShiftTime) < TimDiff(sngShiftTime, sngOne) Then
                    typVar.strShiftOfDay = adrsDumb("shift")
                Else
                    adrsDumb.MovePrevious
                    typVar.strShiftOfDay = adrsDumb("shift")
                End If
            Else
                If TimDiff(sngOne, sngShiftTime) < TimDiff(sngShiftTime, sngTwo) Then
                    adrsDumb.MovePrevious
                    typVar.strShiftOfDay = adrsDumb("shift")
                Else
                    typVar.strShiftOfDay = adrsDumb("shift")
                End If
            End If
        End If
    End If
End If
   Exit Sub
Err_particular:
    Call ShowError("GetAutoShift")
    'Resume Next
End Sub

Public Function InRotShift(ByVal strRotCode As String) As Boolean   '' Checks if the
On Error GoTo ERR_P                                                 '' Shift is Found
Dim strArrPatt() As String                                          '' in the Rotation or
Dim adrsDumb As New ADODB.Recordset
InRotShift = False                                                  '' not
If adrsDumb.State = 1 Then adrsDumb.Close
adrsDumb.Open "Select Pattern from Ro_Shift where SCode='" & strRotCode & "'" _
, ConMain
If (adrsDumb.EOF And adrsDumb.BOF) Then
    Exit Function
Else
    strArrPatt() = Split(adrsDumb("Pattern"), ".")
    Dim bytCntTmp As Integer
    For bytCntTmp = 0 To UBound(strArrPatt)
        If typVar.strShiftOfDay = strArrPatt(bytCntTmp) Then InRotShift = True
    Next
End If
Exit Function
ERR_P:
    Call ShowError("InRotShift")
    InRotShift = False
End Function


Public Sub GetPunchesArray()        '' Gets Array of Punches for two Days
On Error GoTo Err_particular        '' i.e Current Date and Current Date + 1
Dim adrsDumb As New ADODB.Recordset
'OAndM
Dim bytCntTmp As Integer
'Dim bytCntTmp As Integer
typVar.blnFoundPunches = False
If adrsDumb.State = 1 Then adrsDumb.Close

    If bytBackEnd = 2 Then
        adrsDumb.Open "Select distinct Dte,t_punch,unqfld,flg,shift from dailypro where (dte=" & strDTEnc & _
        Format(typDT.dtFrom, "dd/mmm/yyyy") & strDTEnc & " or " & "dte=" & strDTEnc & Format(typDT.dtFrom + 1, "dd/mmm/yyyy") & _
        strDTEnc & ") and EmpCode='" & typVar.strEmpCodeTrn & "' order by Dte,t_punch", ConMain, adOpenStatic
    Else
        adrsDumb.Open "Select distinct Dte,t_punch,unqfld,flg,shift from dailypro where (dte = Convert (DateTime, '" & Format(typDT.dtFrom, "dd/mmm/yyyy") & "') or dte = Convert (DateTime, '" & Format(typDT.dtFrom + 1, "dd/mmm/yyyy") & "' )) and EmpCode='" & typVar.strEmpCodeTrn & "' order by Dte,t_punch", ConMain, adOpenStatic
    End If

''
If (adrsDumb.EOF And adrsDumb.BOF) Then
    Exit Sub
End If


If (typPerm.blnIO Or typPerm.blnDI) Then
    ReDim VarArrPunches(adrsDumb.RecordCount - 1, 5)
Else
    ReDim VarArrPunches(adrsDumb.RecordCount - 1, 4)
End If
''
    'this if conditin ad by  12-sep MIS2007DF015
    If typPerm.blnIO Then
        If adrsDumb("DTE") = typDT.dtFrom + 1 And _
             UCase(adrsDumb("SHIFT")) = "I" Then Exit Sub
    ElseIf typPerm.blnDI Then
           If adrsDumb("DTE") = typDT.dtFrom + 1 And _
             UCase(adrsDumb("SHIFT")) = "DI" Then Exit Sub
        'for check 1st punch is previous punch.
        '=1 then yes
        '=0 then today's punch
        If SetP_TimeFlag(adrsDumb("t_punch")) = 1 Then
            If adrsDumb.RecordCount > 1 Then
                adrsDumb.MoveNext
                'If adrsDumb.RecordCount > 1 Then
                    If adrsDumb("DTE") = typDT.dtFrom + 1 And _
                        (UCase(adrsDumb("SHIFT")) = "DI" Or UCase(adrsDumb("SHIFT")) = "I") Then
                        Exit Sub
                    End If
                ''End If
            End If
        End If
    End If
    If Not (adrsDumb.EOF And adrsDumb.BOF) Then adrsDumb.MoveFirst
    ''
For bytCntTmp = 0 To adrsDumb.RecordCount - 1               '' Punch TIme
    VarArrPunches(bytCntTmp, 0) = Val(Format(adrsDumb("t_punch"), "HH.nn"))
    VarArrPunches(bytCntTmp, 1) = adrsDumb("dte")
    VarArrPunches(bytCntTmp, 2) = adrsDumb("UnqFld")        '' Unique Record Number
    VarArrPunches(bytCntTmp, 3) = 0                         '' Usage Flag
    VarArrPunches(bytCntTmp, 4) = IIf(IsNull(adrsDumb("flg")), 9, adrsDumb("flg")) '' P.Flag
    
    If (typPerm.blnIO Or typPerm.blnDI) Then VarArrPunches(bytCntTmp, 5) = adrsDumb("SHIFT")
    ''
    adrsDumb.MoveNext
Next
typVar.blnFoundPunches = True      '' True if Punches are Found
Exit Sub
Err_particular:
    Debug.Print bytCntTmp
    'Resume Next
    Call ShowError("GetPunchesArray")
End Sub

Public Function SetP_TimeFlag(ByVal sngSP As Single) As Integer    '' Sets the P_Time Usage Flag
On Error GoTo Err_particular                                    '' to the Time
Dim bytCntTmp As Integer, sngP_Time As Single
sngP_Time = GetP_Time(typDT.dtFrom - 1, typVar.strEmpCodeTrn)
sngP_Time = Round(sngP_Time - 24, 2)
If sngSP <= sngP_Time And sngP_Time > 0 Then        '' Previous Punch Should be Greater than 0
    SetP_TimeFlag = 1
Else
    SetP_TimeFlag = 0
End If
Exit Function
Err_particular:
    Call ShowError("SetP_TimeFlag")
End Function

Public Function ValidatePunches() As Integer           '' Validates all the Punches
On Error GoTo Err_particular
'Dim bytCntTmp As integer
'OAndM
Dim bytCntTmp As Integer
Dim bytCntZero As Integer

Dim bytFlgTmp As Integer, dtShift As Date
Dim sngTmp As Single

Dim blnIN As Boolean
''
bytFlgTmp = 0
If typVar.blnFoundPunches = False Then
    bytFlgTmp = 1                                   '' If No Punches are Found
Else
    Call MakeRightpunches           '' Check P Punch
    '' For Auto Shift
    sngTmp = -1
    bytCntZero = 0
    For bytCntTmp = 0 To UBound(VarArrPunches)
        If VarArrPunches(bytCntTmp, 3) = 0 Then
            bytCntZero = bytCntZero + 1
            If sngTmp = -1 Then
            
                If VarArrPunches(bytCntTmp, 0) <= 28 Then
                    sngTmp = VarArrPunches(bytCntTmp, 0) '' First Valid Punch
                    
                    If typPerm.blnDI Then blnIN = IIf(VarArrPunches(bytCntTmp, 5) = "DI", True, False)
                    If typPerm.blnIO Then blnIN = IIf(VarArrPunches(bytCntTmp, 5) = "I", True, False)
                    ''
                End If
            ''
            End If
        End If
    Next

    If sngTmp <> -1 Then
        Select Case typVar.strShiftTmp
            Case typVar.strWosCode, typVar.strHlsCode
                If typEmp.blnAutoOnPunch Then typEmp.blnAuto = True
        End Select
        If typEmp.blnAuto Then
            
            If (typPerm.blnIO Or typPerm.blnDI) Then
                If blnIN Then Call GetAutoShift(sngTmp)
            Else
                Call GetAutoShift(sngTmp)
            End If
            ''
         End If
    End If
          


    Dim blnFirst As Boolean
    Select Case typVar.strShiftTmp
        Case typVar.strWosCode, typVar.strHlsCode
            Call FillShifttype(typVar.strShiftOfDay)
            For bytCntTmp = 0 To UBound(VarArrPunches)      '' Check Upto Time
                If VarArrPunches(bytCntTmp, 0) >= 26 And Not blnFirst Then
                    VarArrPunches(bytCntTmp, 3) = 2
                Else
                    blnFirst = True
                    
                    ''If VarArrPunches(bytCntTmp, 0) > typShift.sngOut + typPerm.sngUpto Then _
                        VarArrPunches(bytCntTmp, 3) = 2
                    If VarArrPunches(bytCntTmp, 0) > typShift.sngOut + typShift.sngUPTO Then _
                        VarArrPunches(bytCntTmp, 3) = 2
                    ''
                End If
            Next
        Case Else
            Call FillShifttype(typVar.strShiftOfDay)
            
            i = 0
            For bytCntTmp = 0 To UBound(VarArrPunches)      '' Check Upto Time
                
                If VarArrPunches(bytCntTmp, 0) > typShift.sngOut + typShift.sngUPTO Then
                     If typPerm.blnIO Then                                    '  If Codition Add By  08-01
                        'If VarArrPunches(bytCntTmp, 1) = typDT.dtFrom + 1 Then
                            VarArrPunches(bytCntTmp, 3) = 2
                        'End If
                     Else
                    VarArrPunches(bytCntTmp, 3) = 2
                    End If
                Else 'this else condition add by  MIS2007DF015
                    If typPerm.blnIO Then
                        If (UCase(VarArrPunches(bytCntTmp, 5)) = "DI" Or UCase(VarArrPunches(bytCntTmp, 5)) = "I") And _
                            VarArrPunches(bytCntTmp, 1) = typDT.dtFrom + 1 Then
                                VarArrPunches(bytCntTmp, 3) = 2
                        End If
                    End If
                End If
            Next
            
    End Select
''
    bytCntZero = 0
    '' Sort out those Punches with the Flag 0 i.e Valid Useable Punch
    For bytCntTmp = 0 To UBound(VarArrPunches)
        If VarArrPunches(bytCntTmp, 3) = 0 Then
            bytCntZero = bytCntZero + 1
        End If
    Next
    typVar.bytTmpEnt = bytCntZero                '' This the No of the Total Valid Punches
    If typVar.bytTmpEnt <= 0 Then
        bytFlgTmp = 2       '' If No Zero Punches Are Available
    Else
        Call GetAflgDflg    '' Gets the Aflg & the Dflg Flags
    End If
End If
ValidatePunches = bytFlgTmp
Exit Function
Err_particular:
    Call ShowError("ValidatePunches")
    ValidatePunches = 1
    'Resume Next
End Function
Public Sub GetLastShift(ByVal strDtShf As String, ByVal strEmp As String)
On Error GoTo Err_particular            '' Get Shift of the Day from the Monthly
Dim adrsDumb As New ADODB.Recordset
Dim strTmp As String                    '' Transaction File
strTmp = Left(MonthName(Month(DateCompDate(strDtShf))), 3) & Right(CStr(Year(DateCompDate(strDtShf))), 2) & "Shf"
If adrsDumb.State = 1 Then adrsDumb.Close
If FindTable(strTmp) Then
    adrsDumb.Open "Select D" & Day(DateCompDate(strDtShf)) & " from " & strTmp & " where Empcode='" & typEmp.strEmp _
    & "'", ConMain
    If Not (adrsDumb.EOF And adrsDumb.BOF) Then
        typVar.strShiftOfDay = adrsDumb(0)
    Else
        typVar.strShiftOfDay = ""
    End If
Else
    typVar.strShiftOfDay = ""
End If
Exit Sub
Err_particular:
    Call ShowError("GetTrnShift")
    typVar.strShiftOfDay = ""
End Sub

Public Sub AddRecordsToTrn(Optional bytFromCorrection As Integer = 0)         '' Adds Records to the Monthly Transaction File
On Error GoTo Err_particular
Dim strTmp As String
'' Create Monthly Transaction File Name
Dim status As String
   strTmp = MonthName(Month(typDT.dtFrom), True) & Right(CStr(Year(typDT.dtFrom)), 2) & "Trn"

'' If Monthly Transaction File not found then Create it
If strTmp <> strFlag And strFlag <> "" Then ' 13-07
    AdFieldChangeSt = False
End If

'Call RoundInOutTime

strFlag = strTmp
If bytFromCorrection = 1 Then blnCurrDtTrnFound = True
If Not blnCurrDtTrnFound Then
    'conmain.Execute " Select * into " & strTmp & " from Montrn"
    Call CreateTableIntoAs("*", "Montrn", strTmp)
    Call CreateTableIndexAs("MONYYTRN", MonthName(Month(typDT.dtFrom), True), Right(CStr(Year(typDT.dtFrom)), 2))
    blnCurrDtTrnFound = True

End If

       ConMain.Execute "Delete from " & strTmp & " Where " & strKDate & "=" & _
       strDTEnc & Format(typDT.dtFrom, "dd/mmm/yyyy") & strDTEnc & " and Empcode='" & typEmp.strEmp & "'"



        Call GetPresent



        ConMain.Execute "insert into " & strTmp & " values('" & _
        typEmp.strEmp & "'," & strDTEnc & Format(typDT.dtFrom, "dd/mmm/yyyy") & strDTEnc & "," & typVar.bytTmpEnt & _
        "," & typEmp.bytEntry & _
        ",'" & typVar.strShiftOfDay & "'," & typTR.sngTimeIn & "," & typDH.sngLateHrs & "," & _
        typTR.sngBreakE & "," & typTR.sngBreakS & "," & typTR.sngBreakHrs & "," & _
        typTR.sngTimeOut & "," & typDH.sngEarlyHrs & "," & typDH.sngWorkHrs & "," & _
        typDH.sngOverTime & "," & typTR.sngTime5 & "," & typTR.sngTime6 & "," & _
        typTR.sngTime7 & "," & typTR.sngTime8 & "," & typTR.sngODFrom & "," & _
        typTR.sngODTo & "," & typTR.sngOFDFrom & "," & typTR.sngOFDTo & "," & typTR.sngOFDHrs & _
        "," & typVar.sngPresent & ",'" & typVar.strStatus & "'," & typDH.sngCOHrs & ",'" & typVar.strIrrMark & "','" & _
        typVar.strAflg & "','" & typVar.strDflg & "',0,0,0,'','" & typVar.strRemarks & "','" & _
        IIf(typDH.sngOverTime > 0, typOTVars.strOTAuth, "") & "','" & _
        IIf(typDH.sngOverTime > 0, typOTVars.strOTRem, "") & "' )"    '& IIf(GetFlagStatus("NOMIS2010"), "", IIf(StatusChangLEHour, ",1", ",0")) & ",0)"
 Exit Sub
Err_particular:
    Call ShowError("AddRecordsToTrn")
    Resume Next
End Sub

Private Sub ChangeEntry()
 Dim bytCount As Integer
    typTR.sngBreakE = 0
    typTR.sngBreakS = 0
    typTR.sngTime5 = 0
    typTR.sngTime6 = 0
    typTR.sngTime7 = 0
    typTR.sngTime8 = 0
    If typVar.bytTmpEnt >= 2 Then
        If typTR.sngTimeIn <> 0 And typTR.sngTimeOut = 0 Then
            typVar.bytTmpEnt = 1
            typDH.sngWorkHrs = TimDiff(typTR.sngTimeOut, typTR.sngTimeIn)
            typVar.strIrrMark = "*"
        ElseIf typTR.sngTimeIn = 0 And typTR.sngTimeOut <> 0 Then
            typVar.bytTmpEnt = 1
            typDH.sngWorkHrs = TimDiff(typTR.sngTimeOut, typTR.sngTimeIn)
            typVar.strIrrMark = "*"
        Else
            typVar.bytTmpEnt = 2
            typDH.sngWorkHrs = TimDiff(typTR.sngTimeOut, typTR.sngTimeIn)
            typVar.strIrrMark = ""
        End If
    End If
End Sub


Private Function GetP_Time(ByVal strDtShf As Date, ByVal strEmp As String) As Single
On Error GoTo Err_particular                '' Get's the P_Time
Dim strTmp As String, sngCounter As Single  '' P_Time is the Last Punch of the Last
                                            '' Day Processed
Dim adrsDumb As New ADODB.Recordset
Dim sngP_Time As Single
strTmp = Left(MonthName(Month(strDtShf)), 3) & _
Right(CStr(Year(strDtShf)), 2) & "trn"
If FindTable(strTmp) Then
    If adrsDumb.State = 1 Then adrsDumb.Close
    adrsDumb.Open "Select deptim,time8,time7,time6,time5,actrt_i,actrt_o,arrtim from " & _
    strTmp & " where Empcode='" & typVar.strEmpCodeTrn _
    & "' and " & strKDate & "=" & strDTEnc & Format(strDtShf, "dd/mmm/yyyy") & strDTEnc, ConMain
    sngCounter = 8
    If Not (adrsDumb.EOF And adrsDumb.BOF) Then
        Do
            Select Case sngCounter
                Case Is = 8
                    sngP_Time = IIf(IsNull(adrsDumb("deptim")), 0, adrsDumb("deptim"))
               Case Is = 7
                    sngP_Time = IIf(IsNull(adrsDumb("time8")), 0, adrsDumb("time8"))
                Case Is = 6
                    sngP_Time = IIf(IsNull(adrsDumb("time7")), 0, adrsDumb("time7"))
                Case Is = 5
                    sngP_Time = IIf(IsNull(adrsDumb("time6")), 0, adrsDumb("time6"))
                Case Is = 4
                    sngP_Time = IIf(IsNull(adrsDumb("time5")), 0, adrsDumb("time5"))
                Case Is = 3
                    sngP_Time = IIf(IsNull(adrsDumb("actrt_i")), 0, adrsDumb("actrt_i"))
                Case Is = 2
                    sngP_Time = IIf(IsNull(adrsDumb("actrt_o")), 0, adrsDumb("actrt_o"))
                Case Is = 1
                    sngP_Time = IIf(IsNull(adrsDumb("arrtim")), 0, adrsDumb("arrtim"))
            End Select
            If sngP_Time <> 0 Then sngCounter = 0
            sngCounter = sngCounter - 1
        Loop Until sngCounter < 1
    Else
        sngP_Time = 0
    End If
Else
    sngP_Time = 0
End If
GetP_Time = sngP_Time
Exit Function
Err_particular:
    Call ShowError("GetP_Time")
    GetP_Time = 0
End Function

Public Sub SettypEmp()      '' Set type Employee
typEmp.blnAuto = False
''typEmp.blnComp = False
''typEmp.blnOT = False
typEmp.bytEntry = 0
typEmp.strCard = ""
typEmp.strECat = ""
typEmp.strConv = ""
typEmp.strEmp = ""
typEmp.strName = ""
typEmp.strOff = ""
typEmp.strOff2 = ""
typEmp.strWO13 = ""
typEmp.strWO24 = ""
typEmp.strEmpShift = ""
typEmp.strShifttype = ""
typOTVars.bytOTCode = 100
typCOVars.bytCOCode = 100
'' For Details regarding Daily Processing
typEmp.bytWOHLAction = 0
typEmp.strAction3Shift = ""
typEmp.blnAutoOnPunch = False
typEmp.strActionBlank = ""

typEmp.strAutoGroup = ""
''
End Sub

Public Sub SettypShift()    '' Set type Shift
typShift.blnNight = False
typShift.sngB1I = 0
typShift.sngB1O = 0
typShift.sngB2I = 0
typShift.sngB2O = 0
typShift.sngB3I = 0
typShift.sngB3O = 0
typShift.sngBH1 = 0
typShift.sngBH2 = 0
typShift.sngBH3 = 0
typShift.sngHalfE = 0
typShift.sngHalfS = 0
typShift.sngHRS = 0
typShift.sngIN = 0
typShift.sngOut = 0
typShift.strShift = ""

typShift.sngUPTO = 0
''
End Sub

Public Sub SettypCat()      '' Set type Category
typCat.sngCutE = 0
typCat.sngCutL = 0
typCat.sngEarl = 0
typCat.sngEarlI = 0
typCat.sngLate = 0
typCat.sngLateI = 0
typCat.strCat = ""
typCat.strDesc = ""

typCat.sngLunchLtIgnore = 0
End Sub

Public Sub SettypTR()       '' Set type Time Range
typTR.sngBreakE = 0
typTR.sngBreakHrs = 0
typTR.sngBreakS = 0
typTR.sngODFrom = 0
typTR.sngODTo = 0
typTR.sngOFDFrom = 0
typTR.sngOFDHrs = 0
typTR.sngOFDTo = 0
typTR.sngTime5 = 0
typTR.sngTime6 = 0
typTR.sngTime7 = 0
typTR.sngTime8 = 0
typTR.sngTimeIn = 0
typTR.sngTimeOut = 0
End Sub

Public Sub SettypDH()       '' Set type Daily Hours
typDH.sngEarlyHrs = 0
typDH.sngLateHrs = 0
typDH.sngRndLateHrs = 0 '15-04
typDH.sngOverTime = 0
typDH.sngWorkHrs = 0
typDH.sngCOHrs = 0
End Sub

Public Sub SettypOTVars()       '' Set type OTvars
'' WD/WO/HL
typOTVars.bytOTHL = 0
typOTVars.bytOTWD = 0
typOTVars.bytOTWO = 0
typOTVars.sngWDRate = 0
typOTVars.sngWORate = 0
typOTVars.sngHLRate = 0
'' Authorization
typOTVars.strOTAuth = ""
'' Max OT
typOTVars.sngMaxOT = 0
'' Late Early
typOTVars.bytDedEarl = 0
typOTVars.bytDedLate = 0
'' Deductions
typOTVars.sngF1 = 0
typOTVars.sngT1 = 0
typOTVars.sngD1 = 0
typOTVars.bytAll1 = 0
typOTVars.sngF2 = 0
typOTVars.sngT2 = 0
typOTVars.sngD2 = 0
typOTVars.bytAll2 = 0
typOTVars.sngF3 = 0
typOTVars.sngT3 = 0
typOTVars.sngD3 = 0
typOTVars.bytAll3 = 0
typOTVars.sngMoreThan = 0
typOTVars.sngD4 = 0
typOTVars.bytAll4 = 0
typOTVars.bytApplyHL = 0
typOTVars.bytApplyWO = 0
'' Round Off
typOTVars.sngRF1 = 0
typOTVars.sngRT1 = 0
typOTVars.sngR1 = 0
typOTVars.sngRF2 = 0
typOTVars.sngRT2 = 0
typOTVars.sngR2 = 0
typOTVars.sngRF3 = 0
typOTVars.sngRT3 = 0
typOTVars.sngR3 = 0
typOTVars.sngRF4 = 0
typOTVars.sngRT4 = 0
typOTVars.sngR4 = 0
typOTVars.sngRT5 = 0
typOTVars.sngR5 = 0
'' Remarks
typOTVars.strOTRem = ""
End Sub

Public Sub SettypCOVars()
'' WD\WO\HL
typCOVars.bytCOWD = 0
typCOVars.bytCOWO = 0
typCOVars.bytCOHL = 0
'' Avail Limit
typCOVars.bytCOAvail = 0
'' CO Calculation Hours
typCOVars.sngWDF = 0
typCOVars.sngWDH = 0
typCOVars.sngWOF = 0
typCOVars.sngWOH = 0
typCOVars.sngHLH = 0
typCOVars.sngHLF = 0
'' Late Early
typCOVars.bytCOLate = 0
typCOVars.bytCOEarl = 0
End Sub

Public Sub SetBreakHours()  '' Set type Break Hours
typBH.sngBrk1 = 0
typBH.sngBrk2 = 0
typBH.sngBrk3 = 0

typBH.sngBrk1Late = 0
End Sub

Public Sub SetMiscVars()    '' Initializes and Re-Initializes Other Variables
typVar.strIrrMark = ""
typVar.strRemarks = ""
typVar.strStatus = ""
typVar.strShiftOfDay = ""
typVar.strShiftTmp = ""
typVar.sngPresent = 1
typVar.bytTmpEnt = 0
typVar.strTmpLvtype = ""
typVar.strAflg = ""
typVar.strDflg = ""
typVar.strLeaveStatus = ""
typVar.blnTrnLate = False
typVar.blnTrnErl = False
End Sub

Private Sub PutTimeRange(Optional bytFlgTR As Integer = 1)     '' Puts the Valid Time Punches
On Error GoTo ERR_P                                         '' in the typTR Created for Time
If Not typVar.blnFoundPunches Then Exit Sub                        '' Ranges
Dim bytCntTR As Integer, bytCntZero As Integer, bytOD As Integer
If typVar.blnFoundPunches = False Then Exit Sub                    '' If no Punches
Dim sngArrTmpZero() As Single
bytCntZero = 0: bytOD = 0
'' Sort out those Punches with the Flag 0 i.e Valid Useable Punch
For bytCntTR = 0 To UBound(VarArrPunches)
    '' Code to be Inserted here for Special Flags
    Select Case VarArrPunches(bytCntTR, 4)
        Case 0          '' OD Flag
            If VarArrPunches(bytCntTR, 3) = 0 Then
                If typShift.blnNight = True Then
                        If typTR.sngODFrom = 0 Then     '' If Night Shift
                            typTR.sngODFrom = VarArrPunches(bytCntTR, 0) + 24
                        Else
                            typTR.sngODTo = VarArrPunches(bytCntTR, 0) + 24
                        End If
                Else                                '' If Normal Shift
                    If typTR.sngODFrom = 0 Then
                        typTR.sngODFrom = VarArrPunches(bytCntTR, 0)
                    Else
                        typTR.sngODTo = VarArrPunches(bytCntTR, 0)
                    End If
                End If
                ''COMMENTED BY SG07
                ''bytOD = bytOD + 1
            End If
        Case Else       '' Normal Flag
            If VarArrPunches(bytCntTR, 3) = 0 Then
                bytCntZero = bytCntZero + 1
                ReDim Preserve sngArrTmpZero(bytCntZero - 1)
                sngArrTmpZero(bytCntZero - 1) = VarArrPunches(bytCntTR, 0)
            End If
    End Select
Next
bytCntZero = bytCntZero
If bytCntZero = 0 Then Exit Sub
'' Put those Punches in the typTR Variable
bytCntZero = UBound(sngArrTmpZero) + 1
typVar.bytTmpEnt = bytCntZero + bytOD
Select Case bytCntZero
    Case 1
        typTR.sngTimeIn = sngArrTmpZero(0)
    Case 2
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngTimeOut = sngArrTmpZero(1)
    Case 3
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngBreakE = sngArrTmpZero(1)
        typTR.sngTimeOut = sngArrTmpZero(2)
    Case 4
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngBreakE = sngArrTmpZero(1)
        typTR.sngBreakS = sngArrTmpZero(2)
        typTR.sngTimeOut = sngArrTmpZero(3)
    Case 5
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngBreakE = sngArrTmpZero(1)
        typTR.sngBreakS = sngArrTmpZero(2)
        typTR.sngTime5 = sngArrTmpZero(3)
        typTR.sngTimeOut = sngArrTmpZero(4)
    Case 6
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngBreakE = sngArrTmpZero(1)
        typTR.sngBreakS = sngArrTmpZero(2)
        typTR.sngTime5 = sngArrTmpZero(3)
        typTR.sngTime6 = sngArrTmpZero(4)
        typTR.sngTimeOut = sngArrTmpZero(5)
    Case 7
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngBreakE = sngArrTmpZero(1)
        typTR.sngBreakS = sngArrTmpZero(2)
        typTR.sngTime5 = sngArrTmpZero(3)
        typTR.sngTime6 = sngArrTmpZero(4)
        typTR.sngTime7 = sngArrTmpZero(5)
        typTR.sngTimeOut = sngArrTmpZero(6)
    Case 8
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngBreakE = sngArrTmpZero(1)
        typTR.sngBreakS = sngArrTmpZero(2)
        typTR.sngTime5 = sngArrTmpZero(3)
        typTR.sngTime6 = sngArrTmpZero(4)
        typTR.sngTime7 = sngArrTmpZero(5)
        typTR.sngTime8 = sngArrTmpZero(6)
        typTR.sngTimeOut = sngArrTmpZero(7)
    Case Is > 8
        typTR.sngTimeIn = sngArrTmpZero(0)
        typTR.sngBreakE = sngArrTmpZero(1)
        typTR.sngBreakS = sngArrTmpZero(2)
        typTR.sngTime5 = sngArrTmpZero(3)
        typTR.sngTime6 = sngArrTmpZero(4)
        typTR.sngTime7 = sngArrTmpZero(5)
        typTR.sngTime8 = sngArrTmpZero(6)
        typTR.sngTimeOut = sngArrTmpZero(UBound(sngArrTmpZero))     '' The Last Valid Punch
        typVar.bytTmpEnt = 8
End Select

 
'' In Case of Unfound Shift or 0 or 1 Entry
  
If typPerm.blnIgnore = True Then
    Call ChangeEntry
End If
Exit Sub
ERR_P:
    ShowError ("Put Time Range :: Daily Process")
End Sub

Private Sub PutTimeRangeIO(Optional bytFlgTR As Integer = 1)     '' Puts the Valid Time Punches
On Error GoTo ERR_P                                         '' in the typTR Created for Time
If Not typVar.blnFoundPunches Then Exit Sub                        '' Ranges
Dim bytCntTR As Integer, bytCntZero As Integer, bytOD As Integer
If typVar.blnFoundPunches = False Then Exit Sub                    '' If no Punches
Dim sngArrTmpZero() As Single
''For Deemah
Dim bytINTmp As Integer, bytOUTTmp As Integer
Dim intTmp As Integer, strIOTmp As String
''
bytOUT = 0: bytIN = 0
bytCntZero = 0: bytOD = 0
'' Sort out those Punches with the Flag 0 i.e Valid Useable Punch
For bytCntTR = 0 To UBound(VarArrPunches)
    '' Code to be Inserted here for Special Flags
    Select Case VarArrPunches(bytCntTR, 4)
        Case 0          '' OD Flag
            If VarArrPunches(bytCntTR, 3) = 0 Then

                If typShift.blnNight = True Then
                    VarArrPunches(bytCntTR, 0) = VarArrPunches(bytCntTR, 0) + 24
                End If
                Select Case bytOD
                    Case 0
                        typTR.sngODFrom = VarArrPunches(bytCntTR, 0)
                    Case 1
                        typTR.sngODTo = VarArrPunches(bytCntTR, 0)
                    Case 2
                        typTR.sngOFDFrom = VarArrPunches(bytCntTR, 0)
                    Case 3
                        typTR.sngOFDTo = VarArrPunches(bytCntTR, 0)
                End Select
                ''
                bytOD = bytOD + 1
            End If
        Case Else       '' Normal Flag
            If VarArrPunches(bytCntTR, 3) = 0 Then
                bytCntZero = bytCntZero + 1
                ReDim Preserve sngArrTmpZero(bytCntZero - 1)
                ReDim Preserve strArrIO(bytCntZero - 1)
                sngArrTmpZero(bytCntZero - 1) = VarArrPunches(bytCntTR, 0)
                strArrIO(bytCntZero - 1) = VarArrPunches(bytCntTR, 5)
                If strArrIO(bytCntZero - 1) = "I" Then
                    bytIN = bytIN + 1
                Else
                    bytOUT = bytOUT + 1
                End If
            End If
    End Select
Next
bytCntZero = bytCntZero
If bytCntZero = 0 Then Exit Sub
'' Put those Punches in the typTR Variable
bytCntZero = UBound(sngArrTmpZero) + 1
typVar.bytTmpEnt = bytCntZero + bytOD
''For deemah
If bytIN > 0 Then bytINTmp = 1
If bytOUT > 0 Then bytOUTTmp = 1

''For Deemah
For intTmp = 0 To UBound(sngArrTmpZero)
    If strArrIO(intTmp) = "I" Then
        Select Case bytINTmp
            Case 1: typTR.sngTimeIn = sngArrTmpZero(intTmp)
            Case 2: typTR.sngBreakS = sngArrTmpZero(intTmp): bytOUTTmp = 2
            Case 3: typTR.sngTime6 = sngArrTmpZero(intTmp): bytOUTTmp = 3
            Case 4: typTR.sngTime8 = sngArrTmpZero(intTmp): bytOUTTmp = 4
            'case else add by  MIS2007DF017
            Case Else:
                If typTR.sngTimeIn = 0 And typTR.sngBreakS = 0 And _
                    typTR.sngTime6 = 0 And typTR.sngTime8 = 0 Then
                    typTR.sngTimeIn = sngArrTmpZero(intTmp)
                End If
            ''
        End Select
        bytINTmp = bytINTmp + 1
        If strIOTmp = "I" Then blnIrregular = True
        strIOTmp = "I"
    Else
        Select Case bytOUTTmp
            Case 1: typTR.sngBreakE = sngArrTmpZero(intTmp): bytINTmp = 2
            Case 2: typTR.sngTime5 = sngArrTmpZero(intTmp): bytINTmp = 3
            Case 3: typTR.sngTime7 = sngArrTmpZero(intTmp): bytINTmp = 4
            Case 4: typTR.sngTimeOut = sngArrTmpZero(intTmp)
            'case else add by  MIS2007DF017
            Case Else:
                typTR.sngTimeOut = sngArrTmpZero(intTmp)
            ''
        End Select
        bytOUTTmp = bytOUTTmp + 1
        If strIOTmp = "O" Then blnIrregular = True
        strIOTmp = "O"
    End If
Next
''
  
''In
If typPerm.blnIgnore = True Then
    If typTR.sngTimeIn > 0 Then
        ''do nothing
    ElseIf typTR.sngBreakS > 0 Then
        typTR.sngTimeIn = typTR.sngBreakS
        typTR.sngBreakS = 0
    ElseIf typTR.sngTime6 > 0 Then
        typTR.sngTimeIn = typTR.sngTime6
        typTR.sngTime6 = 0
    ElseIf typTR.sngTime8 > 0 Then
        typTR.sngTimeIn = typTR.sngTime8
        typTR.sngTime8 = 0
    End If
    ''Out
    If typTR.sngTimeOut > 0 Then
        ''do nothing
    ElseIf typTR.sngTime7 > 0 Then
        typTR.sngTimeOut = typTR.sngTime7
        typTR.sngTime7 = 0
    ElseIf typTR.sngTime5 > 0 Then
        typTR.sngTimeOut = typTR.sngTime5
        typTR.sngTime5 = 0
    ElseIf typTR.sngBreakE > 0 Then
        typTR.sngTimeOut = typTR.sngBreakE
        typTR.sngBreakE = 0
    End If
    If typTR.sngTimeIn <> 0 And typTR.sngTimeOut <> 0 Then
        blnIrregular = False
    Else
        blnIrregular = True
    End If
    Call ChangeEntry
    Exit Sub
End If
''
Dim blnIO As Boolean
Dim sngTmp As Single, bytFor As Byte
For bytFor = 1 To 8
    Select Case bytFor
        Case 1
            sngTmp = typTR.sngTimeIn
            blnIO = True
        Case 2
            sngTmp = typTR.sngBreakE
            blnIO = False
        Case 3
            sngTmp = typTR.sngBreakS
            blnIO = True
        Case 4
            sngTmp = typTR.sngTime5
            blnIO = False
        Case 5
            sngTmp = typTR.sngTime6
            blnIO = True
        Case 6
            sngTmp = typTR.sngTime7
            blnIO = False
        Case 7
            sngTmp = typTR.sngTime8
            blnIO = True
        Case 8
            sngTmp = typTR.sngTimeOut
            blnIO = False
    End Select
    If sngTmp = 0 And blnIO = True Then
        Select Case bytFor
            Case 1
                blnIrregular = True
                typTR.sngTimeOut = typTR.sngBreakE
                typTR.sngBreakE = 0
                Exit For
            Case 3
                If typTR.sngTime5 = 0 Then
                    typTR.sngTimeOut = typTR.sngBreakE
                    typTR.sngBreakE = 0
                    Exit For
                End If
            Case 5
                If typTR.sngTime7 = 0 Then
                    typTR.sngTimeOut = typTR.sngTime5
                    typTR.sngTime5 = 0
                    Exit For
                End If
            Case 7
                If typTR.sngTimeOut = 0 Then
                    typTR.sngTimeOut = typTR.sngTime7
                    typTR.sngTime7 = 0
                    Exit For
                End If
        End Select
    End If
    If sngTmp = 0 And blnIO = False Then
        blnIrregular = True
        Exit For
    End If
Next
If bytIN <> bytOUT Then blnIrregular = True
Exit Sub
ERR_P:
    ShowError ("Put Time Range IO:: Daily Process")
    'Resume Next
End Sub
''
Public Sub NoShift()        '' Action to be Taken if no Shift is Found for Processing
On Error GoTo ERR_P
Call MakeRightpunches       '' Make Right Punches

If typPerm.blnIO Then
    Call PutTimeRangeIO(2)
Else
    Call PutTimeRange(2)        '' Get Time Hours
End If
''
Call GetStatus(1)              '' Get Status
Select Case typEmp.bytEntry
    Case 0          '' Provision for 0 Entry
    Case 1          '' Provision for 1 Entry
    Case Else
        Call GetWorkHours(9)        '' Get Work Hours
        Call GetLateHours(3)        '' Check Late Hours
        Call GetEarlyHours(3)       '' Check Early Hours
End Select
Call AddRecordsToTrn
'' Add Records to the TRN File
Exit Sub
ERR_P:
    ShowError ("NoShift")
End Sub

Public Function GetLeaveStatus()            '' Calculates and Returns the LeaveStatus for
On Error GoTo ERR_P                         '' that Date
Dim dt1 As Date, dt2 As Date, blnFlg As Boolean
Dim strTmp As String, strLvTmp As String
Dim adrsDumb As New ADODB.Recordset

' on 08/03/05
If typPerm.intYrFrom = 4 Then
    If Month(typDT.dtFrom) < 4 Then
        strTmp = "LvInfo" & Right(Year(typDT.dtFrom) - 1, 2)
    Else
        strTmp = "LvInfo" & Right(Year(typDT.dtFrom), 2)
    End If
Else
    strTmp = "LvInfo" & Right(Year(typDT.dtFrom), 2)
End If

'strTmp = "LvInfo" & Right(Year(typDT.dtFrom), 2)
If FindTable(strTmp) Then
    If adrsDumb.State = 1 Then adrsDumb.Close
    adrsDumb.Open "Select * from " & strTmp & " where Empcode=" & "'" & typEmp.strEmp & _
    "'" & " and (" & strDTEnc & Format(typDT.dtFrom, "dd/mmm/yyyy") & strDTEnc & _
    " between fromdate and todate ) and trcd=4", ConMain, adOpenStatic
    If Not (adrsDumb.BOF And adrsDumb.EOF) Then
        GetLeaveStatus = "": strLvTmp = ""
        Do While Not adrsDumb.EOF
            dt1 = adrsDumb!FromDate: dt2 = adrsDumb!ToDate
            If DateCompDate(typDT.dtFrom) = adrsDumb("fromdate") And adrsDumb("hf_option") = "FST " Then
                GetLeaveStatus = "  " & adrsDumb!LCode
            ElseIf DateCompDate(typDT.dtFrom) = adrsDumb!ToDate And adrsDumb("hf_option") = "F TF" Then
                GetLeaveStatus = adrsDumb!LCode & "  "
            ElseIf (DateCompDate(typDT.dtFrom) = adrsDumb!FromDate And DateCompDate(typDT.dtFrom) = adrsDumb!ToDate) And adrsDumb("hf_option") = "FSTS" Then
                GetLeaveStatus = "  " & adrsDumb!LCode
            ElseIf (DateCompDate(typDT.dtFrom) = adrsDumb!FromDate And DateCompDate(typDT.dtFrom) = adrsDumb!ToDate) And adrsDumb("hf_option") = "FFTF" Then
                GetLeaveStatus = adrsDumb!LCode & "  "
            Else
                GetLeaveStatus = adrsDumb!LCode & "" & adrsDumb!LCode
            End If
            GetLeaveStatus = strLvTmp & GetLeaveStatus
            If Len(GetLeaveStatus) > 4 Then GetLeaveStatus = Replace(GetLeaveStatus, " ", "")
            typVar.strTmpLvtype = adrsDumb!Lv_Type_rw
            adrsDumb.MoveNext
            If Not adrsDumb.EOF Then
                If dt1 = adrsDumb!FromDate And dt2 = adrsDumb!ToDate Then strLvTmp = GetLeaveStatus
            End If
        Loop
    End If
    adrsDumb.Close
Else
    GetLeaveStatus = ""
End If
Exit Function
ERR_P:
    ShowError ("GetLeaveStatus")
End Function


Public Sub GetAflgDflg()
Dim bytCntTmp As Integer                   '' Counter to Array
'OAndM
'Dim bytCntTmp As Integer
Dim sngPermFlg() As Single       '' Permission Flag Array
Dim bytCntFlg As Integer                   '' Counter of Zero Record Punches
If UBound(VarArrPunches) >= 7 Then
    ReDim sngPermFlg(UBound(VarArrPunches) + 1) As Single
Else
    ReDim sngPermFlg(8) As Single
End If
bytCntFlg = 1
For bytCntTmp = 0 To UBound(VarArrPunches)
    If VarArrPunches(bytCntTmp, 3) = 0 Then
        sngPermFlg(bytCntFlg) = VarArrPunches(bytCntTmp, 4)
        bytCntFlg = bytCntFlg + 1
    End If
Next
If (sngPermFlg(1) = "1" Or sngPermFlg(1) = "3") Then typVar.strAflg = sngPermFlg(1)
If (sngPermFlg(1) = "4") Then typVar.strDflg = "4"
If (sngPermFlg(3) = "1" Or sngPermFlg(3) = "3") Then typVar.strAflg = sngPermFlg(3)
If (sngPermFlg(5) = "1" Or sngPermFlg(5) = "3") Then typVar.strAflg = sngPermFlg(5)
If (sngPermFlg(7) = "1" Or sngPermFlg(7) = "3") Then typVar.strAflg = sngPermFlg(7)
If (sngPermFlg(2) = "2") = "2" Then typVar.strDflg = sngPermFlg(2)

If (sngPermFlg(4) = "2") = "2" Then typVar.strDflg = sngPermFlg(4)
If (sngPermFlg(6) = "2") = "2" Then typVar.strDflg = sngPermFlg(6)
If (sngPermFlg(8) = "2") = "2" Then typVar.strDflg = sngPermFlg(8)
If UBound(VarArrPunches) >= 7 Then
    For i = 8 To UBound(VarArrPunches) + 1
        If sngPermFlg(i) = "2" Then typVar.strDflg = sngPermFlg(i)
    Next
End If

End Sub

Public Sub GetShiftTimings()        '' Special Case when Shift is O Shift
If typShift.strShift = "O" Then typShift.sngIN = typTR.sngTimeIn
End Sub

Public Sub GetWorkHours(Optional bytR2 As Integer = 1)
On Error GoTo ERR_P
Dim sngTmp As Single

If typPerm.blnIO And blnIrregular Then Exit Sub
''
If typShift.strShift = "O" Then bytR2 = 9 'OPEN SHIFT

Select Case bytR2
    Case 1      '' Get the Basic Work Hours
        If typVar.strShiftTmp = typVar.strHlsCode Or typVar.strShiftTmp = typVar.strWosCode Then
            typDH.sngWorkHrs = TimDiff(typTR.sngTimeOut, typTR.sngTimeIn)
        Else
                typDH.sngWorkHrs = typShift.sngHRS
        End If
        'End If
    Case 2      '' Actual Work Hours After Basic Late and Early Hours
        If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode Then
            If typDH.sngLateHrs > 0 And typVar.strAflg <> "1" And typVar.strAflg <> "3" Then typDH.sngWorkHrs = TimDiff(typDH.sngWorkHrs, typDH.sngLateHrs)
            If typDH.sngEarlyHrs > 0 And typVar.strDflg <> "2" Then typDH.sngWorkHrs = TimDiff(typDH.sngWorkHrs, typDH.sngEarlyHrs)
        End If
        
    Case 3      '' Actual Work Hours after the Late Allowance is Calculated

            If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode And _
            typDH.sngLateHrs <= 0 Then
                '' Changes as on 21-05-2003
                typDH.sngWorkHrs = TimAdd(typDH.sngWorkHrs, Abs(typDH.sngLateHrs))
                'End If
                Call GetStatus(4)
                ''
                End If
         
    Case 4      '' Actual Work Hours after the Early Allowance is Calculated
        If typVar.strShiftTmp <> typVar.strHlsCode And typVar.strShiftTmp <> typVar.strWosCode And _
        typDH.sngEarlyHrs <= 0 Then
            '' Changes as on 21-05-2003
            typDH.sngWorkHrs = TimAdd(typDH.sngWorkHrs, Abs(typDH.sngEarlyHrs))
            If typDH.sngWorkHrs <= 0 Then typDH.sngWorkHrs = 0
            'End If
            Call GetStatus(6)
            ''
        End If
    Case 5      '' Actual Work Hours after the Break Hours are Calculated
            
 
           sngTmp = TimAdd(typShift.sngBH1, TimAdd(typShift.sngBH2, typShift.sngBH3))
    
            If typTR.sngBreakHrs > sngTmp Then
                sngTmp = TimDiff(typTR.sngBreakHrs, sngTmp)
                typDH.sngWorkHrs = TimDiff(typDH.sngWorkHrs, sngTmp)
            End If
            'End If
    Case 6      '' When Odd Entries are Foud
            '' If typVar.bytTmpEnt > 0 Then typDH.sngWorkHrs = TimDiff(typTR.sngTimeOut, typTR.sngTimeIn)
    Case 7      '' Set the Work Hours to Zero
        typDH.sngWorkHrs = 0
    Case 8
        If typTR.sngTimeIn > 0 Then
            typDH.sngWorkHrs = typShift.sngHRS
        Else
            typDH.sngWorkHrs = 0
        End If
    Case 9      '' For Work Hours when no shift is Found
        If typTR.sngTimeOut <> 0 Then _
        typDH.sngWorkHrs = TimDiff(typTR.sngTimeOut, typTR.sngTimeIn)
End Select
    
Exit Sub
ERR_P:
    ShowError ("GetWorkHours")
End Sub

''''''''''''''
Public Sub GetHoursZeroEnt(Optional bytFromCorrection As Integer = 0)         '' Procedure When Zero Entry is Required
On Error GoTo ERR_P
GetWorkHours (8)   '' Workhours
'' Check for any Single punch
GetPresent          '' Get Present"
Call AddRecordsToTrn(bytFromCorrection)    '' Add Record
Exit Sub
ERR_P:
    ShowError ("GetHoursZeroEnt")
End Sub

Public Sub GetHoursOneEnt()     '' Procedure When Single Entry is Required
On Error GoTo ERR_P
If typVar.strShiftTmp <> typVar.strWosCode And typVar.strShiftTmp <> typVar.strHlsCode Then
    If typTR.sngTimeIn > 0 Then
        If typDH.sngLateHrs < typPerm.sngPostEarl Or typDH.sngLateHrs < typPerm.sngPostLt Then
            If Left(typVar.strStatus, 2) = "  " Or Left(typVar.strStatus, 2) = typVar.strAbsCode Then
                Call GetStatus(9)
            End If
            If Right(typVar.strStatus, 2) = "  " Or Right(typVar.strStatus, 2) = typVar.strAbsCode Then
                Call GetStatus(10)
            End If
        End If
    Else
        Call GetStatus(11)
        typVar.bytTmpEnt = 0
    End If
End If
Call GetWorkHours(8)    '' Calculate Work Hours
Call GetPresent          '' Get Present"
Call AddRecordsToTrn    '' Add Record
Exit Sub
ERR_P:
    ShowError ("GetHoursOneEnt")
End Sub
Public Sub ProcessHours()
On Error GoTo ERR_P

DoEvents: DoEvents
''
Call GetShiftTimings
Call GetWorkHours                   '' Get Basic Work Hours
Call GetLateHours                   '' Get Basic Late Hours
Call GetEarlyHours                  '' Get Basic Early Hours
Call GetStatus(2)                   '' **   Repeated Phase Calls
Call GetBreak1                      '' Get the First Break Hours
Call GetBreak2                      '' Get the Second Break Hours
Call GetBreak3                      '' Get the Third Break Hours
Call GetTotalBreakHour              '' Get the Total Break Hours
Call GetWorkHours(2)                '' **   Repeated Phase Calls
Call GetWorkHours(3)                '' **   Repeated Phase Calls
Call GetWorkHours(4)                '' **
Call GetWorkHours(5)                '' **
Call GetStatus(7)                   '' **
Call GetStatus(8)                   '' **
Call GetWorkHours(6)                '' **   Repeated Phase Calls
Call GetStatus(12)                  '' Minimum Work Hours Check
Call GetOvertimeHours               '' Get the Basic OT Hours
Call GetCOHrs                       '' Get the Basic CO days
Call GetRemarks                     '' Get the Remarks
Call GetIrrMark                     '' Get the Irregular Mark
Call GetPresent                     '' Get the Present Depending upon the Status


'''''''''''''''''''
Exit Sub
ERR_P:
    ShowError ("Process Hours :: Daily Process")
End Sub

Private Sub RoundInOutTime()
    Dim TR As New clsTimes
    Dim TLeft, TRight, OLeft, ORight
    TRight = Right(Format(typTR.sngTimeIn, "0.00"), 2) 'Girish
    TLeft = Int(typTR.sngTimeIn)
    ORight = Right(Format(typTR.sngTimeOut, "0.00"), 2)
    OLeft = Int(typTR.sngTimeOut)
    
    
    If TRight > 0 And TRight < 31 Then
        TR.sngTimeIn = TimAdd(typTR.sngTimeIn, (30 - TRight) / 100)
    ElseIf TRight > 30 And TRight <= 59 Then
            TR.sngTimeIn = TimAdd(typTR.sngTimeIn, (60 - TRight) / 100)
    Else
            TR.sngTimeIn = TLeft
    End If
    
    If ORight >= 0 And ORight < 30 Then
        TR.sngTimeOut = TimAdd(typTR.sngTimeOut, -(ORight / 100))
    ElseIf ORight >= 30 And ORight <= 59 Then
            TR.sngTimeOut = TimAdd(typTR.sngTimeOut, -(ORight - 30) / 100)
    Else
            TR.sngTimeOut = OLeft
    End If
            
    typDH.sngWorkHrs = TimDiff(TR.sngTimeOut, TR.sngTimeIn)
    If (typDH.sngWorkHrs) > 8.3 Then
        typDH.sngOverTime = TimDiff(typDH.sngWorkHrs, 8.3)
        typDH.sngWorkHrs = 8.3
    End If
    
End Sub

Public Sub PutHours()       '' Put Late , Early & Comp Off Hours

DoEvents: DoEvents
''
Call PutLate
Call PutEarly
Call PutCOHrs
End Sub
''For Mauritius 14-08-2003
Private Sub FillMissedCards()
Dim intCnt As Integer
StrGroup1 = ""
Dim adrsDumb As New ADODB.Recordset
adrsDumb.Open "Select distinct Card from DailyPro Where (Empcode is null or Empcode='') and ( Card not in ('" & typPerm.strOD & _
"','" & typPerm.strLate & "','" & typPerm.strEarl & "','" & typPerm.strBus & "') or Card is not Null)", ConMain, adOpenStatic, adLockOptimistic
Do While Not adrsDumb.EOF
    If intCnt Mod 20 = 0 Then
        StrGroup1 = StrGroup1 & "," & IIf(IsNull(adrsDumb("Card")), "", adrsDumb("Card"))
    Else
        StrGroup1 = StrGroup1 & IIf(IsNull(adrsDumb("Card")), "", adrsDumb("Card")) & vbCrLf
    End If
    adrsDumb.MoveNext
Loop
End Sub
''
Public Sub MakeRightpunches()           '' Make Right Punches
On Error GoTo ERR_P
Dim bytCntTmp As Integer
'OAndM
'Dim bytCntTmp As Integer
If typVar.blnFoundPunches = False Then Exit Sub
For bytCntTmp = 0 To UBound(VarArrPunches)
    If VarArrPunches(bytCntTmp, 1) = typDT.dtFrom Then
        '' Check if the Punch is Processed in any Previous Day
        VarArrPunches(bytCntTmp, 3) = SetP_TimeFlag(VarArrPunches(bytCntTmp, 0))
    End If
    '' For Normal and Next Day Punches Add 24
    If VarArrPunches(bytCntTmp, 3) = 0 And _
        VarArrPunches(bytCntTmp, 1) <> typDT.dtFrom Then _
    VarArrPunches(bytCntTmp, 0) = VarArrPunches(bytCntTmp, 0) + 24
Next
Exit Sub
ERR_P:
    ShowError ("MakeRightPunches :: Daily Module")
End Sub


