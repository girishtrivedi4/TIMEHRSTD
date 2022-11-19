Attribute VB_Name = "CommBas"
Option Explicit

Dim XlsFileName As String
Dim fso As New FileSystemObject
Dim wkbnew As Excel.Workbooks
Public MultiDB As Boolean
Public ConLog As New ADODB.Connection
Public ConMain As New ADODB.Connection

Public Sub Main()
On Error GoTo Err_particular
Dim strTmp1 As String, strTmp2 As String
blnShowLang = False '' Multi Language Option
'' Check For Instance

'true for text datatype and false for numeric datatype for dept
'Test Of SVN
blnFlagForDept = False

strVersionWithTital = GetVersion(True)
strVersionWithOutTital = GetVersion(False)

blnGoWithExportInExcell = False
Call GetValueForTag

    NewCapFlag = False

    SubLeaveFlag = 0
'If Not FreeSpace Then
'    MsgBox "Atlease 100 MB Free Disk Space Required::Cannot Proceed", vbCritical, strVersionWithTital
'    End
'End If
'' Check Initiazation File
If Not ReadIni Then
    MsgBox "Initialization File Error :: Cannot Proceed", vbCritical, strVersionWithTital
    End
End If
If Not ConnectLog Then
    MsgBox "Error Connecting to the Local Database :: Cannot Proceed" & _
    vbCrLf & "Contact Your System Administrator", vbCritical, GetVersion(True)
    End
End If

    frmDSN.Show vbModal
    If ConMain.ConnectionString = "" Then End
    GoTo ConnectLink
If connect <> "" Then
    frmDSN.Show vbModal
  
    DoEvents
    If connect <> "" Then
        MsgBox "Error Connecting to the Database :: Cannot Proceed" & _
        vbCrLf & "Contact Your System Administrator", vbCritical, strVersionWithTital
        End
    Else
ConnectLink:
        adrsMod.ActiveConnection = ConMain
        adrsMod.CursorType = adOpenStatic
        adrsMod.Open "Select * from NewCaptions Where ID Like 'M%' or ID Like '00%'"
    End If
Else
    ''Unload frmSplash
    adrsMod.ActiveConnection = ConMain
    adrsMod.CursorType = adOpenStatic
    adrsMod.Open "Select * from NewCaptions Where ID Like 'M%' or ID Like '00%'"
End If
'' Try Connecting to the Local MDB File


DoEvents
'' Connections OK
'' Check For Multilingual
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select * from Exc", ConMain, adOpenStatic
If adrsTemp.EOF Then
    strCapField = "CaptEng"
Else
    If IsNull(adrsTemp("Lang")) Then
        strCapField = "CaptEng"
    Else
        strCapField = adrsTemp("Lang")
    End If
End If
'' Set Values
'' Back End, Date Enclosures
Select Case InVar.strSer
    Case "1"    '' MS-SQL Server
        strDTEnc = "'"
        bytBackEnd = 1
        strKDate = "[Date]"
        strKGroup = "[Group]"
        strKDesc = "[Desc]"
        strKOff = "[Off]"
        StrKConcat = " + "
        
        strName = "[Name]"
        strLeft = "LEFT"
        strRight = "RIGHT"
    Case "2"    '' MS-Access
        strDTEnc = "#"
        bytBackEnd = 2
        strKDate = "[Date]"
        strKGroup = "[Group]"
        strKDesc = "[Desc]"
        strKOff = "[Off]"
        StrKConcat = " + "
        
        strName = "[Name]"
        strLeft = "LEFT"
        strRight = "RIGHT"
    Case "3"    '' Oracle
        strDTEnc = "'"
        bytBackEnd = 3
        strKDate = """Date"""
        strKGroup = """Group"""
        strKDesc = """Desc"""
        strKOff = "OFF"
        StrKConcat = " || "
        
        strName = "Name"
        strLeft = "LPAD"
        strRight = "RPAD"
End Select

''
If RecordCnt("UserAccs") = 0 Then       '' Check if Records Exists or not
 
    frmFirst.Show                       '' First User
    Exit Sub
Else
    Call GetDateF
    If bytDateF = 0 Then
        MsgBox NewCaptionTxt("M1001", adrsMod), vbCritical
        End
    Else
    
       If Not DateSettings(bytDateF) Then
            ''Msgbox "Unable To Set the Application Date Settings:: Cannot Proceed", vbCritical, App.EXEName
            End
        End If
        
        Call SetDateFormatSQL

        MainForm.Show
        MainForm.Enabled = False
      
       frmLogin.Show vbModal
       If MainForm.Enabled = False Then Call MainForm.MDIForm_Unload(1)
    End If
End If
Exit Sub
Err_particular:
    ShowError ("Main :: Common")
'    Resume Next
    End
End Sub


Public Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
On Error GoTo ERR_P
Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long
Dim R As Long, V$

RetVal$ = ""
R = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)
If R <> ERROR_SUCCESS Then GoTo Quit_Now

SZ = 256: V$ = String$(SZ, 0)
R = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal V$, SZ)
If R = ERROR_SUCCESS And dwType = REG_SZ Then
    If SZ > 1 Then RetVal$ = Left$(V$, SZ - 1)
Else
    RetVal$ = "--Not String--"
End If
If hInKey = 0 Then R = RegCloseKey(hSubKey)
Quit_Now:
RegGetString$ = RetVal$
Exit Function
ERR_P:
    ShowError ("RegGetString::")
End Function
Private Function FreeSpace() As Boolean
On Error GoTo ERR_P
Dim lRet As Long, cBA As Currency, cBOD As Currency, cBF As Currency, m As Double
lRet = GetDiskFreeSpaceEx(Left(App.path, 3), cBA, cBOD, cBF)
cBF = cBF / 102.4
If cBF < 100 Then Exit Function
FreeSpace = True
Exit Function
ERR_P:
    ShowError ("FreeSpace::")
End Function

Public Function TimDiff(ByVal Val1!, ByVal Val2!) As Single
    TimDiff = (Round(dec2Hrs(hrs2Dec(Val1) - hrs2Dec(Val2)), 2))
End Function

Private Function RecordCnt(ByVal tblName As String)
On Error GoTo Err_particular
Dim adRecRs As New ADODB.Recordset
adRecRs.Open "select count(*) from " & tblName, ConMain, adOpenKeyset, adLockReadOnly
RecordCnt = adRecRs(0)
Set adRecRs = Nothing
Exit Function
Err_particular:
    ShowError ("RecordCnt::")
    Set adRecRs = Nothing
End Function

Public Function CatCode(ByVal CatName As String)
On Error GoTo Err_particular
If AdrsCat.State = 1 Then AdrsCat.Close
    AdrsCat.Open "Select * from Catdesc where " & strKDesc & "=" & "'" & CatName & _
    "'", ConMain, adOpenKeyset, adLockOptimistic
If Not (AdrsCat.BOF And AdrsCat.EOF) Then
    CatCode = AdrsCat(0)
End If
Exit Function
Err_particular:
    ShowError ("Category Code")
End Function

Public Function FindTable(ByVal strName As String) As Boolean
On Error GoTo Err_particular
Dim adrsFT As New ADODB.Recordset
Select Case bytBackEnd
    Case 1  '' SQL Server
        adrsFT.Open "Select count(*) from SysObjects where Name='" & strName & "'" _
        , ConMain
    Case 2  '' Ms-Access
        adrsFT.Open "Select count(*) from MSysObjects where Name='" & strName & "'" _
        , ConMain
    Case 3  ''Oracle
        adrsFT.Open "select count(*) from Tabs where UPPER(Table_Name) ='" & UCase(strName) & "'", _
        ConMain
End Select
If IsEmpty(adrsFT(0)) Or IsNull(adrsFT(0)) Or adrsFT(0) <= 0 Then
    FindTable = False
Else
    FindTable = True
End If
adrsFT.Close
Exit Function
Err_particular:
    ShowError ("FindTable::")
End Function

Public Function FieldExists(ByVal tblName$, ByVal FieldName$) As Boolean
'' CHECKS  IF A FIELD EXISTS IN THE SELECTED TABLE
On Error GoTo Err_particular
Dim intCnt As Integer
If adrsRits.State = 1 Then adrsRits.Close
adrsRits.Open "select * from " & tblName, ConMain, adOpenStatic, adLockOptimistic
For intCnt = 0 To adrsRits.Fields.Count - 1
    If UCase(adrsRits.Fields(intCnt).name) = UCase(FieldName) Then
        FieldExists = True
        Exit For
    End If
Next
Exit Function
Err_particular:
    ShowError ("FieldExists::")
End Function

Public Sub ComboFill(ctlCombo As MSForms.ComboBox, ByVal fillType As Integer, _
ByVal ColCnt As Integer, Optional strDeptTmp As String, Optional mon As String, Optional Year As String)
On Error GoTo Err_particular
Dim EmpArray() As String

Dim cnt%
ctlCombo.clear

If fillType > 20 And fillType < 30 Or fillType = 16 Then
            ctlCombo.ListWidth = "6 cm"
            ctlCombo.ColumnWidths = "2 cm;4.5 cm"
            ctlCombo.ColumnCount = 2
Else
    Select Case ColCnt
        Case 2:
            ctlCombo.ListWidth = "6 cm"
            ctlCombo.ColumnWidths = "4.5 cm;2 cm"
            ctlCombo.ColumnCount = 2
        Case 3:
            ctlCombo.ListWidth = "7 cm"
            ctlCombo.ColumnWidths = "2 cm;4 cm,2 cm"
            ctlCombo.ColumnCount = 3
    End Select
End If


'' Code to Change Fill Type depending on Current Department Selected
If strCurrentUserType = HOD Then
    Select Case fillType
        Case 1: fillType = 14
        Case 2: fillType = 15
        Case 7: fillType = 16
    End Select
End If

Dim strTempforCF As String
Select Case fillType
    Case 1:
        strTempforCF = "select Name, Empcode from empmst order by Empcode"               'Empcode,name
        If GetFlagStatus("LocationRights") And strCurrentUserType <> ADMIN Then  '-28-01-10
            strTempforCF = "select Empcode,name from empmst " & strCurrData & " order by Empcode"
        End If
        
    Case 2: strTempforCF = "select " & strKDesc & ", dept from deptdesc order by dept"                        'dept,desc
    Case 3:
        If CatFlag = True Then
            strTempforCF = "select * from catdesc where cat <> '100' order by cat"                              'cat,desc
        Else
            strTempforCF = "select " & strKDesc & ", cat from catdesc where cat <> '100' order by cat"
        End If
    Case 4: '
            
    Case 5: strTempforCF = "select Cname, Company from company Order by Company"          'code,comp name
    Case 6: strTempforCF = "select Leave, distinct(lvcode)  from Leavdesc where lvcode not in (" & _
            "'" & pVStar.PrsCode & "'" & "," & "'" & pVStar.AbsCode & "'" & "," & "'" & pVStar.WosCode & "'" & "," & "'" & _
            pVStar.HlsCode & "'" & ")" & " and type='Y'" & " and cat=" & "'" & strSql & "'"
    Case 7: strTempforCF = "select Empcode,name,email_id  from empmst order by Empcode"
    Case 8: strTempforCF = "select grupdesc, " & strKGroup & " from groupmst order by " & strKGroup & ""
    Case 9: strTempforCF = "select OTDesc, OTCode from OTRul WHERE OTCode <> 100  order by OTCode"
    Case 10: strTempforCF = "select CODesc, COCode from CORul WHERE COCode <> 100  order by COCode"
    Case 11: strTempforCF = "select LocDesc, Location from Location Order by Location"
    Case 12
            strDeptTmp = EncloseQuotes(strDeptTmp)

                Select Case UCase(Trim(strDeptTmp))
                    Case "", "ALL"
                    strTempforCF = "select name,Empcode from empmst order by Empcode"
                    
                    Case Else
                        ''Original-> strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & _
                        intDeptTmp & " order by Empcode"                               'Empcode,name
                        If strCurrentUserType = HOD Then
                            strTempforCF = "select name, Empcode from empmst " & strCurrData & " and Empmst." & SELCRIT & _
                            " = " & strDeptTmp & " order by Empcode"
                        Else
                            strTempforCF = "select name, Empcode from empmst Where empmst." & SELCRIT & _
                            " = " & strDeptTmp & " order by Empcode"
                        End If
                       
                End Select
                If GetFlagStatus("LocationRights") And strCurrentUserType <> ADMIN Then     '-28-01-10
                    If UCase(Trim(strDeptTmp)) = "ALL" Then
                        strTempforCF = "select name, Empcode from empmst " & strCurrData & " order by Empcode"
                    Else
                        strTempforCF = "select name, Empcode from empmst Where Empmst.Dept = " & strDeptTmp & " And Location In (" & UserLocations & ") order by Empcode"
                    End If
                End If
                


    Case 13: strTempforCF = "select DivDesc, Div from Division Order by Div"
    '' New Cases
    Case 14
        ''Original ->strTempforCF = "select Empcode,name from empmst Where Dept=" & intCurrDept & " order by Empcode"
        strTempforCF = "select name, Empcode  from empmst " & strCurrData & " order by Empcode"
    Case 15
        ''Original-->strTempforCF = "select dept," & strKDesc & " from deptdesc Where Dept=" & intCurrDept & " order by dept"
        strTempforCF = "select distinct deptdesc." & strKDesc & ", deptdesc.dept from empmst,deptdesc where " & _
        "empmst.dept=deptdesc.dept " & Replace(strCurrDept, "'", "")
    Case 16
        ''Original ->strTempforCF = "select Empcode,name,email_id  from empmst Where Dept=" & intCurrDept & " order by Empcode"
        strTempforCF = "select Empcode,name,email_id  from empmst " & strCurrData & " order by Empcode"
    'add  by  for left employee MIS2007DF05
    Case 17
        strTempforCF = "SELECT Name, Empcode FROM empmst WHERE leavdate IS NULL ORDER BY empcode"
    Case 18
       
    Case 19
        strTempforCF = "SELECT Empcode, Name FROM empmst WHERE leavdate IS NULL ORDER BY empcode"
        ctlCombo.ListWidth = "6 cm"
        ctlCombo.ColumnWidths = "2 cm; 5 cm"
        ctlCombo.ColumnCount = 2
    Case 20
        strTempforCF = "SELECT designame, desigcode FROM frmDesignation order by desigcode, designame"
            ctlCombo.ListWidth = "6 cm"
            ctlCombo.ColumnWidths = "4.5 cm;2 cm"
            ctlCombo.ColumnCount = 2
    Case 21
        strTempforCF = "select Company, Cname  from company Order by Company"
    Case 22
        strTempforCF = "select dept," & strKDesc & "  from deptdesc order by dept"
    Case 23
        strTempforCF = "select cat, " & strKDesc & "  from catdesc where cat <> '100' order by cat"
    Case 24
        strTempforCF = "select " & strKGroup & ", grupdesc  from groupmst order by " & strKGroup & ""
    Case 25
        strTempforCF = "select Location, LocDesc  from Location Order by Location"
    Case 26
        strTempforCF = "select Div, DivDesc  from Division Order by Div"
    Case 27
        strTempforCF = "select OTCode, OTDesc from OTRul WHERE OTCode <> 100  order by OTCode"
    Case 28
        strTempforCF = "select COCode, CODesc from CORul WHERE COCode <> 100  order by COCode"
End Select
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strTempforCF, ConMain, adOpenKeyset, adLockBatchOptimistic, adCmdText
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    cnt = adrsTemp.RecordCount
    Select Case ColCnt
        Case 2:
            ReDim EmpArray(cnt% - 1, 1)
            For i = 0 To cnt - 1
                EmpArray(i, 0) = adrsTemp(0)
                EmpArray(i, 1) = adrsTemp(1)
                adrsTemp.MoveNext
            Next i
        Case 3:
            ReDim EmpArray(cnt% - 1, 2)
            For i = 0 To cnt - 1
                EmpArray(i, 0) = adrsTemp(0)
                EmpArray(i, 1) = adrsTemp(1)
                EmpArray(i, 2) = IIf(IsNull(adrsTemp(2)), "", adrsTemp(2))
                adrsTemp.MoveNext
            Next i
     End Select
    ctlCombo.List = EmpArray
End If
adrsTemp.Close
Erase EmpArray
Exit Sub
Err_particular:
    ShowError ("ComboFill::")
'    Resume Next
End Sub

Public Function ReplicateVal(ByVal rStr$, ByVal cntVal%)
If rStr = pVStar.AbsCode Or rStr = pVStar.PrsCode Then
    ReplicateVal = rStr & rStr 'rose
Else
    ReplicateVal = rStr & rStr
End If
End Function

Public Function StuffVal(ByVal actStr$, ByVal startPos%, ByVal RemStr%, ByVal RepStr$)
If startPos = 1 Then
    If Len(RepStr) = 1 Then
        StuffVal = RepStr & Space(1) & Right(actStr, 2)
    ElseIf Len(RepStr) = 2 Then
        StuffVal = RepStr & Right(actStr, 2)
    End If
ElseIf startPos = 3 Then
    StuffVal = Left(actStr, 2) & RepStr
End If
End Function

Public Function con_HrMin(ByVal tm As Single) As Single
Dim fracTim%, intVal%
intVal = Fix(tm + 0.1)
fracTim = (tm - intVal) * 100
con_HrMin = Fix(tm + 0.1) + Fix(fracTim / 60) + (fracTim Mod 60) / 100
End Function

Public Function TimAdd(ByVal arg1 As Single, ByVal arg2 As Single) As Single
    TimAdd = (Round(dec2Hrs(hrs2Dec(arg1) + hrs2Dec(arg2)), 2))
End Function

Public Function dec2Hrs(ByVal ar1 As Single) As Single
Dim SgnFg As Boolean
If ar1 < 0 Then
    SgnFg = True
    ar1 = Abs(ar1)
End If
dec2Hrs = (Round((Fix(ar1) + (ar1 - Fix(ar1)) * 60 / 100), 2))
If SgnFg Then dec2Hrs = dec2Hrs * -1
SgnFg = False
End Function

Public Function hrs2Dec(ByVal ar1 As Single) As Single
    hrs2Dec = (Round((Fix(ar1) + (ar1 - Fix(ar1)) * 100 / 60), 2))
End Function

Public Function Year_Start(Optional bytMonth As Byte = 0, Optional intYear As Integer = 0) As Date
On Error GoTo Err_particular
If bytMonth = 0 Then bytMonth = Val(pVStar.Yearstart)
If intYear = 0 Then intYear = Right(pVStar.YearSel, 2)
Year_Start = "01-" & Left(MonthName(bytMonth), 3) & "-" & CStr(intYear)
Exit Function
Err_particular:
    ShowError ("InstallYearStart::")
End Function

Public Function keycheck(ByVal KeyAscii As Integer, txt As Object) As Integer  'For Keycheck
On Error Resume Next
Dim strKey As String
If KeyAscii <> 8 Then
    If Len(txt.Text) = txt.SelLength Then
        strKey = "1234567890."
        GoTo KeyPut
    End If
    If Val(txt.Text) > 99.99 Then
        strKey = "."
        GoTo KeyPut
    End If
    If InStr(txt.Text, ".") Then
        strKey = "1234567890"
        If Len(txt.Text) > (InStr(txt.Text, ".") + 1) Then strKey = ""
    Else
        If Len(txt.Text) < 2 Then
            strKey = "1234567890."
        ElseIf Len(txt.Text) = 2 Then
            txt.Text = txt.Text & "."
            SendKeys "{END}", True
            strKey = "1234567890"
        ElseIf Len(txt.Text) > 2 Then
            strKey = "1234567890"
        End If
    End If
KeyPut:
    If (InStr(strKey, Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = KeyAscii
    End If
End If
keycheck = KeyAscii
End Function

Public Sub GF(txt As Object)
txt.SelStart = 0
txt.SelLength = Len(txt.Text)
End Sub

Public Function strYearFrom(ByVal dt)
On Error GoTo Err_particular
If Month(dt) < Val(pVStar.Yearstart) Then
    strYearFrom = Year(dt) - 1
Else
    strYearFrom = Year(dt)
End If
Exit Function
Err_particular:
    ShowError ("strYearFrom::")
End Function

Public Sub SetDailyFlag()       '' Procedure to set the Daily Flag in Case of Errors
On Error GoTo Err_particular
ConMain.Execute "Update Exc set Daily=0"
Exit Sub
Err_particular:
    ShowError ("Set Daily Flag::")
End Sub

Public Sub SetMonthlyFlag()     '' Procedure to set the Monthly Flag in Case of Errors
On Error GoTo Err_particular
ConMain.Execute "Update Exc set Monthly=0"
Exit Sub
Err_particular:
    ShowError ("Set Monthly Flag::")
End Sub

Public Sub SetYearlyFlag(ByVal bytFlg As Byte)       '' Procedure to set the Yearly Leaves Flag in Case of Errors
On Error GoTo Err_particular
ConMain.Execute "Update Exc set Yearly=" & bytFlg
Exit Sub
Err_particular:
    ShowError ("Set Yearly Flag::")
End Sub

Public Sub ShowError(ByVal strFunction As String)       '' Shows Error Messages 21/07/2001
    MsgBox strFunction & vbCrLf & Err.Description & vbCrLf & "Operation Cancelled", _
    vbCritical, strVersionWithTital

    Call WriteLog(strFunction & vbCrLf & Err.Description & ":" & Erl)
    'Resume Next
End Sub

Public Sub RetUserNumber()              '' Gets the UserNumber of the Userlogged in
On Error GoTo Err_particular
    intUserNum = 0
    If adrsRits.State = 1 Then adrsRits.Close
    adrsRits.Open "Select * from Exc", ConMain
    intUserNum = adrsRits("UserNumber")
    If intUserNum >= 1000 Then ConMain.Execute "Update Exc set UserNumber=0"
Exit Sub
Err_particular:
    ShowError ("Return User Number::")
End Sub

Public Sub WriteLog(strLOGG As String)
On Error Resume Next
If Dir(App.path & "\VSLOG.LOG") = "" Then
    Open App.path & "\VSLOG.LOG" For Output As #2
    Print #2, IIf(UCase(UserName) = UCase(strPrintUser) And UserName <> "" _
    , "*****", UserName) & " :: " & Now & vbCrLf & strLOGG
    Close #2
Else
    Open App.path & "\VSLOG.LOG" For Append As #2
    Print #2, "----------------------" & vbCrLf & IIf(UCase(UserName) = _
    UCase(strPrintUser) And UserName <> "", "*****", UserName) & " :: " & Now & _
    vbCrLf & strLOGG
    Close #2
End If
End Sub

Public Function EncryptSCR(ByVal pwd As String) As String
Dim EncrptStr As String
For i = 1 To Len(pwd)
    EncrptStr = EncrptStr & Chr((Asc(Mid(pwd, i, 1)) Xor 9) Xor 2)
Next i
EncryptSCR = EncrptStr
End Function

Public Function DecryptSCR(ByVal pwd As String) As String
Dim pwdStr As String
For i = 1 To Len(pwd)
    pwdStr = pwdStr & Chr((Asc(Mid(pwd, i, 1)) Xor 9) Xor 2)
Next i
DecryptSCR = pwdStr
End Function

Public Sub TruncateTable(ByVal strTabName As String)    '' Truncates the Table based on the
On Error GoTo ERR_P                                     '' Backend
Select Case bytBackEnd
    Case 1      '' Default SQL Server
        ConMain.Execute "Truncate table " & strTabName
    Case 2      '' MS-Access
        ConMain.Execute "Delete From  " & strTabName
    Case 3      '' Oracle
        ConMain.Execute "Truncate table " & strTabName
End Select
Exit Sub
ERR_P:
    ShowError ("Truncate Table :: DJCommon")
End Sub

Public Sub GetDateF()           '' Gets the Type of the Date Format
On Error GoTo ERR_P
'' Start Code to Get the the Application Date Format
If adRsInstall.State = 1 Then adRsInstall.Close
adRsInstall.Open "Select american_dt from install", ConMain
If (adRsInstall.EOF And adRsInstall.BOF) Then
    bytDateF = 0
Else
    If adRsInstall("american_dt") = False Then
        bytDateF = 2        '' British
    Else
        bytDateF = 1        '' American
    End If
End If
Exit Sub
ERR_P:
    ShowError ("GetDateF::")
End Sub

Public Function KeyPressCheck(ByVal KeyAscii As Integer, Optional bytKey As Byte = 1) As Integer
If KeyAscii <> 8 Then           '' Checks the Keypress Events of the Specified Edit Boxes
    Dim strKeyAscii As String
    Select Case bytKey
        Case 1  '' Only Characters
            strKeyAscii = "qazwsxedcrfvtgbyhnujmikolpQAZWSXEDCRFVTGBYHNUJMIKOLP"
        Case 2  '' Only Numbers
            strKeyAscii = "1234567890"
        Case 3  '' Characters & Space
            strKeyAscii = "qazwsxedcrfvtgbyhnujmikolpQAZWSXEDCRFVTGBYHNUJMIKOLP "
        Case 4  '' Numbers & Decimal
            strKeyAscii = "1234567890."
        Case 5  '' Characters and Numbers With Space
            strKeyAscii = "qazwsxedcrfvtgbyhnujmikolpQAZWSXEDCRFVTGBYHNUJMIKOLP 1234567890"
        Case 6  '' Characters and Numbers Without Space
            strKeyAscii = "qazwsxedcrfvtgbyhnujmikolpQAZWSXEDCRFVTGBYHNUJMIKOLP1234567890"
        Case 7  '' Characters,Numbers,Spaces and Email Characters
            strKeyAscii = "qazwsxedcrfvtgbyhnujmikolpQAZWSXEDCRFVTGBYHNUJMIKOLP1234567890 " & _
                          "@_-."
        Case 8  '' Adddress Characters
            strKeyAscii = "qazwsxedcrfvtgbyhnujmikolpQAZWSXEDCRFVTGBYHNUJMIKOLP1234567890 " & _
                          "#,.@/\+-"
    End Select
    If bytKey = 2 Or bytKey = 4 Then
        If InStr(strKeyAscii, Chr(KeyAscii)) <= 0 Then KeyAscii = 0
    Else
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    End If
    'If InStr(strKeyAscii, Chr(KeyAscii)) <= 0 Then KeyAscii = 0
End If
KeyPressCheck = KeyAscii
End Function

Public Sub SetDateFormatSQL()               '' If SQL Server and the Date Type is 2 then
On Error GoTo ERR_P                         '' Execute the DMY Statements
If bytBackEnd = 1 And bytDateF = 2 Then
    ConMain.Execute "Set DateFormat DMY"
End If
If bytBackEnd = 3 Then
    Select Case bytDateF
        Case 1 ''American
            ConMain.Execute "ALTER SESSION SET NLS_DATE_FORMAT ='MM/dd/yy'"
        Case 2 ''British
             ConMain.Execute "ALTER SESSION SET NLS_DATE_FORMAT ='dd/mm/yy'"
     End Select
    
End If
Exit Sub
ERR_P:
    ShowError ("Set SQL Server Date Format::")
End Sub

Public Sub SetFormIcon(ByRef frm As Object, Optional blnTmpFlg As Boolean = False) '' Sets the Form Icon
On Error GoTo ERR_P
'frm.Icon = LoadPicture(App.Path & "\Images\Starsys.Ico")
frm.Icon = LoadPicture(App.path & "\Images\TimeHR.ico")
If blnTmpFlg Then Exit Sub
AddRights = False
EditRights = False
DeleteRights = False
Exit Sub
ERR_P:
    ShowError ("Set Form Icon::")
End Sub

Public Function KeyDecimal3(ByVal KeyAscii As Integer, txt As Object) As Integer  'For KeyDecimal3
On Error Resume Next
Dim strKey As String
If KeyAscii <> 8 Then
    If Len(txt.Text) = txt.SelLength Then
        strKey = "1234567890."
        GoTo KeyPut
    End If
    If Val(txt.Text) > 999.99 Then
        strKey = "."
        GoTo KeyPut
    End If
    If InStr(txt.Text, ".") Then
        strKey = "1234567890"
        If Len(txt.Text) > (InStr(txt.Text, ".") + 1) Then strKey = ""
    Else
        If Len(txt.Text) < 3 Then
            strKey = "1234567890."
        ElseIf Len(txt.Text) = 3 Then
            txt.Text = txt.Text & "."
            SendKeys "{END}", True
            strKey = "1234567890"
        ElseIf Len(txt.Text) > 3 Then
            strKey = "1234567890"
        End If
    End If
KeyPut:
    If (InStr(strKey, Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = KeyAscii
    End If
End If
KeyDecimal3 = KeyAscii
End Function

Public Sub SetToolTipText(ByRef frm As Form)        '' Sets the ToolTip Text for the Date
On Error GoTo ERR_P                                 '' TextBoxes. Set the Value of Tag
Dim cntrl1 As Control                               '' Property to 'D' for the TextBoxes
For Each cntrl1 In frm.Controls                     '' for Dates
    If TypeOf cntrl1 Is TextBox Then
        If cntrl1.Tag = "D" Then cntrl1.ToolTipText = strDateFO
    End If
Next
Exit Sub
ERR_P:
    ShowError ("Set ToolTip Text::")
End Sub

Public Sub SetCaptionMainForm()
On Error GoTo ERR_P
If UCase(Trim(UserName)) <> strPrintUser Then

    'MainForm.Caption = UCase(App.EXEName) & Space(2) & "2011" & _
                        Space(5) & "User Name : " & UserName & Space(5) & InVar.strCOM & _
                        Space(1) & IIf(InVar.blnVerType = "1", NewCaptionTxt("M1015", adrsMod), "")
    MainForm.Caption = "User Name : " & UserName & Space(5) & InVar.strCOM & _
                        Space(1) '& IIf(InVar.blnVerType = "1", NewCaptionTxt("M1015", adrsMod), "")
Else
    'MainForm.Caption = UCase(App.EXEName) & Space(2) & "2011" & _
                        Space(5) & "User Name : " & "******" & Space(5) & InVar.strCOM & _
                        Space(1) & IIf(InVar.blnVerType = "1", NewCaptionTxt("M1015", adrsMod), "")
    MainForm.Caption = IIf(InVar.blnVerType = "1", UCase(App.EXEName) & Space(2) & NewCaptionTxt("M1015", adrsMod), UCase(App.EXEName) & Space(2)) & _
                        Space(5) & MainForm.Caption = "User Name : " & "******" & Space(5) & InVar.strCOM & _
                        Space(1)  '& IIf(InVar.blnVerType = "1", NewCaptionTxt("M1015", adrsMod), "")
End If
''
  
MainForm.Caption = MainForm.Caption & Space(50) & strVersionWithTital
Exit Sub
ERR_P:
    ShowError ("SetCaptionMainForm::")
End Sub

Public Function ReadIni() As Boolean
On Error GoTo ERR_P
Dim strArrDec() As String, strArrTmp() As String, strData As String
strData = ""
Open App.path & "\Data\TimeHR.ini" For Binary As #2
strData = Trim(Input(LOF(2), #2))
Close #2
strArrTmp = Split(strData, vbCrLf)
strArrTmp(1) = DEncryptDat(strArrTmp(1), 2)
If strArrTmp(0) = strArrTmp(1) Then
    strArrTmp(0) = DEncryptDat(strArrTmp(0), 1)
    strArrDec = Split(strArrTmp(0), "|")
    '' Company Name
    InVar.strCOM = strArrDec(0)
    '' Demo Version
    InVar.blnVerType = strArrDec(1)
    '' Employee Limit
    InVar.lngEmp = strArrDec(2)
    '' Net Type
    InVar.blnNetType = strArrDec(3)
    '' User Limit
    InVar.bytUse = strArrDec(4)
    '' Company Limit
    InVar.bytCom = strArrDec(5)
    '' Version Number
    InVar.strVer = strArrDec(6)
    '' Lock Date
    InVar.strLok = strArrDec(7)
    '' Assum
    InVar.blnAssum = 0 ' strArrDec(8)
    '' Back End
    InVar.strSer = strArrDec(9)
    '' Web Enabled
    InVar.blnWeb = strArrDec(10)
    '' User Name
    InVar.strUser = strArrDec(11)
    '' Password
    InVar.strPass = strArrDec(12)
    ''Location
'    InVar.strLoc = strArrDec(13)
Else
    ShowError ("ReadIni::FileTampering")
    Exit Function
End If
ReadIni = True
Exit Function
ERR_P:
    If Err.Number = 9 Then
        MsgBox "Invalid INI file, please contact IV SOFTTECH", vbCritical, GetVersion(True)
    Else
        ShowError ("ReadIni::")
    End If
    ''Resume Next
End Function

Public Function DEncryptDat(pwd As String, bytTimes As Byte) As String
Dim pwdStr As String, i As Integer
For i = 1 To Len(pwd)
        pwdStr = pwdStr & Chr(Asc(Mid(pwd, i, 1)) Xor IIf(bytTimes = 1, 11, 12))
Next i
DEncryptDat = pwdStr
End Function

Public Function connect()
On Error GoTo ERR_P
Dim strTmp1 As String, strTmp2 As String
Dim strTmp3 As String
'' Get Registry Settings
'' User Name
strTmp1 = RegGetString(HKEY_LOCAL_MACHINE, "SoftWare\ODBC\ODBC.INI\" & TDSN.DSNName & "", "UID")
'' Pasword
strTmp2 = RegGetString(HKEY_LOCAL_MACHINE, "SoftWare\ODBC\ODBC.INI\" & TDSN.DSNName & "", "PWD")
''Service Name in case of ORACLE
If InVar.strSer = 3 Then
    strTmp3 = RegGetString(HKEY_LOCAL_MACHINE, "SoftWare\ODBC\ODBC.INI\VisualStarDSN", "SERVER")
    strDBPass = "; PWD=" & DEncryptDat(strTmp2, 1)
End If
'For Crystal Report.
StrCrUser = strTmp1
StrcrPwd = strTmp2
StrcrSvr = strTmp3
'************************
'' Try to Connnect
Select Case InVar.strSer
    Case "1"    '' MS-SQL Server
        ConMain.Open "DSN=" & TDSN.DSNName & "; Uid=" & DEncryptDat(strTmp1, 1) & _
        "; PWD=" & DEncryptDat(strTmp2, 1) & ";Database=" & TDSN.Database
          ConMain.CommandTimeout = 600
    Case "2"    '' Ms-Access
    ConMain.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Attendo-2011\Attendodb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ATTENDO"
        ConMain.Open
        strBackEndPath = RegGetString(HKEY_LOCAL_MACHINE, _
        "SoftWare\ODBC\ODBC.INI\VisualStarDSN", "DBQ")
    Case "3"    ''Oracle
        ConMain.Open "Provider=MSDASQL.1;Persist Security Info=False;" & _
        "User ID=" & DEncryptDat(strTmp1, 1) & ";Password=" & DEncryptDat(strTmp2, 1) & _
        ";SERVER=" & strTmp3 & ";Data Source=VisualStarDSN"

        ConMain.Open "Provider=MSDataShape.1;Persist Security Info=False;" & _
        "Data Source=VisualStarDSN;User ID=" & DEncryptDat(strTmp1, 1) & _
        ";Password=" & DEncryptDat(strTmp2, 1) & _
        ";SERVER=" & strTmp3 & ";Data Provider=MSDASQL"

        ConMain.CursorLocation = adUseClient
End Select
Exit Function
ERR_P:
    WriteLog ("Connect::")
    connect = "ERROR_CONNECTING"
End Function

Public Function DateSettings(bytTmp As Byte) As Boolean
On Error GoTo ERR_P
Dim dwLCID As Long
Dim lBuffSize As String
Dim sBuffer As String
Dim lRet As Long
lBuffSize = 256
sBuffer = String$(lBuffSize, vbNullChar)
lRet = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_SSHORTDATE, sBuffer, lBuffSize)
If lRet > 0 Then
    sBuffer = Left$(sBuffer, lRet - 1)
End If
Select Case bytTmp
    Case 1      '' MM-DD-YY
        Select Case UCase(sBuffer)
            Case "M/D/YY", "M/D/YYYY", "MM/DD/YY", "MM/DD/YYYY"
                strDateFO = UCase(sBuffer)
            Case Else
                bytTmp = 10
                MsgBox NewCaptionTxt("M1002", adrsMod) & vbCrLf & _
                NewCaptionTxt("M1003", adrsMod) & _
                NewCaptionTxt("M1004", adrsMod), vbCritical, strVersionWithTital
                Exit Function
        End Select
    Case 2      '' DD-MM-YY
        Select Case UCase(sBuffer)
            Case "D/M/YY", "D/M/YYYY", "DD/MM/YY", "DD/MM/YYYY"
                strDateFO = UCase(sBuffer)
            Case Else
                MsgBox NewCaptionTxt("M1005", adrsMod) & vbCrLf & _
                NewCaptionTxt("M1003", adrsMod) & _
                NewCaptionTxt("M1004", adrsMod), vbCritical, strVersionWithTital
                bytTmp = 20
                Exit Function
        End Select
End Select
Select Case bytTmp
    Case 10     '' MM-D-YY
        If MsgBox(NewCaptionTxt("M1006", adrsMod) & _
        NewCaptionTxt("M1007", adrsMod) & vbCrLf & NewCaptionTxt("M1008", adrsMod), vbQuestion + vbYesNo, strVersionWithTital) _
        = vbYes Then
            dwLCID = GetSystemDefaultLCID()
            If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "M/D/YY") Then
                strDateFO = "M/D/YY"
                DateSettings = True
                PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
                PostMessage HWND_BROADCAST, WM_WININICHANGE, 0, 0
                Exit Function
            End If
        End If
    Case 20     '' DD-MM-YY
        If MsgBox(NewCaptionTxt("M1006", adrsMod) & _
        NewCaptionTxt("M1009", adrsMod) & vbCrLf & NewCaptionTxt("M1008", adrsMod), vbQuestion + vbYesNo, strVersionWithTital) _
        = vbYes Then
            dwLCID = GetSystemDefaultLCID()
            If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "DD/MM/YY") Then
                strDateFO = "DD/MM/YY"
                DateSettings = True
                PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
                PostMessage HWND_BROADCAST, WM_WININICHANGE, 0, 0
                Exit Function
            End If
        End If
    Case Else
        DateSettings = True
        Exit Function
End Select
Exit Function
ERR_P:
    ShowError ("DateSettings::")
End Function
''

Public Sub SetCritLabel(lblObj As Object)
Select Case UCase(SELCRIT)
    Case "DEPT"
        lblObj.Caption = NewCaptionTxt("00058", adrsMod)                '' Department
    Case "CAT"
        lblObj.Caption = NewCaptionTxt("00051", adrsMod)              '' Category
    Case "COMPANY"
        lblObj.Caption = NewCaptionTxt("00057", adrsMod)              '' Company
    Case "" & strKGroup & ""
        lblObj.Caption = NewCaptionTxt("00059", adrsMod)              '' Group
    Case "LOCATION"
        lblObj.Caption = NewCaptionTxt("00110", adrsMod)              '' Location
End Select
End Sub

Public Function EncloseQuotes(ByVal strTmp As String) As String
'' On Error resume Next

    Select Case UCase(SELCRIT)
        Case "DEPT":
        Case "CAT"
            If strTmp <> "ALL" Then strTmp = "'" & strTmp & "'"
        Case "COMPANY"
        Case "" & strKGroup & ""
        Case "LOCATION"
    End Select

EncloseQuotes = strTmp
End Function

Public Sub SetCritCombos(ByRef cboDept As Object)
'' On Error Resume Next
Select Case UCase(SELCRIT)
    Case "DEPT"
       Call ComboFill(cboDept, 2, 2)
    Case "CAT"
        Call ComboFill(cboDept, 3, 2)
    Case "COMPANY"
        Call ComboFill(cboDept, 5, 2)
    Case "" & strKGroup & ""
        Call ComboFill(cboDept, 8, 2)
    Case "LOCATION"
        Call ComboFill(cboDept, 11, 2)
    Case "DIV"
        Call ComboFill(cboDept, 13, 2)
End Select

If strCurrentUserType <> HOD Then cboDept.AddItem "ALL", 0

End Sub

Public Sub SelUnsel(clrObj As ColorConstants, objGrid As Object, cb1 As Object, cb2 As Object) '' Function to Select Unselect the
On Error Resume Next
Dim tempvar                                         '' Employee in the Grid
Dim bytTmp As Byte
If cb1.Text = "" Then Exit Sub
If cb2.Text = "" Then Exit Sub
tempvar = cb2.ListIndex - cb1.ListIndex
Select Case tempvar
    Case Is < 0
        objGrid.Col = 0
        For i = cb2.ListIndex + 1 To cb1.ListIndex + 1
            With objGrid
                .row = i
                For bytTmp = 0 To objGrid.Cols - 1
                    .Col = bytTmp
                    .CellBackColor = clrObj
                Next
                .Col = 0
            End With
        Next i
    Case Is = 0
        objGrid.Col = 0
        With objGrid
            .row = cb1.ListIndex + 1
            For bytTmp = 0 To objGrid.Cols - 1
                .Col = bytTmp
                .CellBackColor = clrObj
            Next
            .Col = 0
        End With
    Case Is > 0
        objGrid.Col = 0
        For i = cb1.ListIndex + 1 To cb2.ListIndex + 1
            With objGrid
                .row = i
                 For bytTmp = 0 To objGrid.Cols - 1
                    .Col = bytTmp
                    .CellBackColor = clrObj
                Next
                .Col = 0
            End With
        Next i
End Select
End Sub

Public Sub SelUnselAll(clrObj As ColorConstants, objGrid As Object)  '' Function to Select UnSelect All the
Dim bytTmp As Byte
For i = 1 To objGrid.Rows - 1                          '' Employees
    With objGrid
        .row = i
        For bytTmp = 0 To objGrid.Cols - 1
            .Col = bytTmp
            .CellBackColor = clrObj
        Next
        .Col = 0
    End With
Next i
End Sub

Public Sub SetGridDetails(frm As Object, fr As Object, MSF As Object, _
lblFrom As Object, lblTo As Object)
On Error GoTo ERR_P
'' Frame Label
fr.Caption = NewCaptionTxt("00044", adrsMod)
'' From & To Label
lblFrom.Caption = NewCaptionTxt("00045", adrsMod)
lblTo.Caption = NewCaptionTxt("00046", adrsMod)
'' Grid Captions
MSF.TextMatrix(0, 0) = NewCaptionTxt("00047", adrsMod)
MSF.TextMatrix(0, 1) = NewCaptionTxt("00048", adrsMod)
'' ToolTip Text
MSF.ToolTipText = NewCaptionTxt("00111", adrsMod)
'' Buttons Label
frm.cmdSR.Caption = "Select Range"
frm.cmdUR.Caption = "Unselect Range"
frm.cmdSA.Caption = "Select All"
frm.cmdUA.Caption = "Unselect All"
Exit Sub
ERR_P:
    ShowError ("SetGridDetails::Common")
End Sub

Public Sub ShowCalendar()
On Error Resume Next
    frmCal.Show vbModal
End Sub

Public Sub SetGButtonCap(frm As Form, Optional bytFlgCap As Byte = 1)   '' Sets Captions to the Main
If bytFlgCap = 1 Then                                       '' Buttons
    frm.cmdAddSave.Caption = "Add"
    frm.cmdEditCan.Caption = "Update"
    frm.cmdDel.Caption = "Delete"
    frm.cmdExit.Caption = "Exit"
Else
    frm.cmdAddSave.Caption = "Save"
    frm.cmdEditCan.Caption = "Cancel"
End If
End Sub

Public Function RetRights(ByVal bytRightType As Byte, ByVal bytRightPos As Byte, _
Optional bytHODRights As Byte = 0, Optional bytCount As Byte = 3) As String
'' On Error Resume Next
Dim strTmp As String
strTmp = "000"
Select Case strCurrentUserType
    Case ADMIN
        strTmp = String(bytCount, "1")
    Case HOD
        Select Case bytRightType
            Case 1          '' Master
                If bytHODRights = 0 Then
                    strTmp = "000"
                Else
                    strTmp = String(3, Mid(strHODRights, bytHODRights, 1))
                End If
            Case 2, 4       '' HOD / Mixed
                strTmp = Mid(strHODRights, bytHODRights, 1)
                If strTmp = "1" Then
                    strTmp = String(bytCount, "1")
                Else
                    strTmp = String(bytCount, "0")
                End If
            Case 3          '' Other
                strTmp = String(bytCount, "0")
        End Select
    Case GENERAL
        Select Case bytRightType
            Case 1  '' Master
                strTmp = Mid(strMasterRights, (bytRightPos * 3) - 2, 3)
            Case 2  '' Leave
                strTmp = Mid(strLeaveRights, (bytRightPos * 2) - 1, 2)
            Case 3  '' Other
                strTmp = Mid(strOtherRights1, bytRightPos, bytCount)
            Case 4  '' Mixed
                strTmp = Mid(strOtherRights1, bytRightPos, bytCount)
        End Select
End Select
RetRights = strTmp
End Function

Public Sub CreateTableIntoAs(ByVal strFldNames As String, ByVal strFromTable As String, ByVal strToTable As String, Optional ByVal str1E2 As String = "")
On Error GoTo ERR_P
Select Case bytBackEnd
    Case 1, 2 ''SQL-SERVER,ACCESS
        ConMain.Execute "select " & strFldNames & " into " & strToTable & " from " & strFromTable & str1E2
    Case 3    ''Oracle
        ConMain.Execute "Create Table " & strToTable & " as select " & strFldNames & " from " & strFromTable & str1E2
End Select
Exit Sub
ERR_P:
'    ShowError ("CreateTableIntoAs :: " & strFldNames & "::" & strFromTable & "::" & strToTable)
    Resume Next
End Sub

Public Function RightStr(ByVal strTmp As String)
Select Case bytBackEnd
    Case 1
        RightStr = "Right(" & strTmp & ",2)"
    Case 2
        RightStr = "Right(" & strTmp & ",2)"
    Case 3
        RightStr = "Substr(" & strTmp & ",3,2)"
End Select
End Function

Public Function LeftStr(ByVal strTmp As String)
Select Case bytBackEnd
    Case 1
        LeftStr = "Left(" & strTmp & ",2)"
    Case 2
        LeftStr = "Left(" & strTmp & ",2)"
    Case 3
        LeftStr = "Substr(" & strTmp & ",1,2)"
End Select
End Function

Public Function AlreadyGiven(strTTable As String, strTField As String, strValue As String) As Boolean
On Error GoTo Err_particular
Dim adrsDumb As New ADODB.Recordset
adrsDumb.Open "Select count(*) from " & strTTable & " where " & strTField & " = " & _
strValue, ConMain, adOpenStatic, adLockOptimistic
If Not adrsDumb.EOF Then
    If adrsDumb(0) <> 0 Then
        MsgBox NewCaptionTxt("00128", adrsMod), vbCritical
        AlreadyGiven = False
    Else
        AlreadyGiven = True
    End If
Else
    AlreadyGiven = True
End If
Exit Function
Err_particular:
    ShowError ("AlreadyGiven :: Combas")
End Function

Public Sub CreateTableIndexAs(ByVal strTblNameTemplate As String, Optional ByVal strMonth As String, Optional ByVal strYear As String)
On Error GoTo ERR_P
Dim strTableName As String
Dim strIndexName As String, strKeyName As String
Dim adRecRs As New ADODB.Recordset
strTblNameTemplate = UCase(strTblNameTemplate)
adRecRs.Open "Select * from tblIndexMaster " & _
             "Where cTblName like '" & strTblNameTemplate & "%'", _
             ConMain, adOpenKeyset, adLockReadOnly
    Do While Not adRecRs.EOF
            If strTblNameTemplate = "MONYYTRN" Then 'Monthly TRN File
                strIndexName = strMonth & strYear & "TRN"
                strTableName = strIndexName
            ElseIf strTblNameTemplate = "MONYYSHF" Then 'Monthly Shift File
                strIndexName = strMonth & strYear & "SHF"
                strTableName = strIndexName
            ElseIf strTblNameTemplate = "LVTRNYY" Then 'Yearly TRN File
                strIndexName = "LVTRN" & strYear
                strTableName = strIndexName
            ElseIf strTblNameTemplate = "LVINFOYY" Then 'Yearly Info File
                strIndexName = "LVINFO" & strYear
                strTableName = strIndexName
            ElseIf strTblNameTemplate = "LVBALYY" Then 'Yearly Balance File
                strIndexName = "LVBAL" & strYear
                strTableName = strIndexName
            Else
                strIndexName = adRecRs.Fields("cIndexName")
                strTableName = adRecRs.Fields("cTblName")
            End If
                strKeyName = Trim(adRecRs.Fields("Cindexkeys"))
             If InStr(1, strKeyName, UCase(strKDate)) <> 0 Then
               strKeyName = Replace(strKeyName, UCase(strKDate), strKDate)
             End If
            
            ConMain.Execute _
            "Create " & _
            Trim(adRecRs.Fields("cUnique")) & _
            " Index " & strIndexName & _
            " ON " & strTableName & "(" & strKeyName & ")"
            adRecRs.MoveNext
    Loop
Exit Sub
ERR_P:
    If Err.Number = -2147217900 Then Resume Next         '        02-03
    ShowError ("CreateTableIndexAs :: " & strTableName & "::" & strIndexName)
    'Resume Next
End Sub
  
Public Function ExportIntoFile(strFileName As String, strQ As String) As String
10    On Error GoTo ERR_P
20        XlsFileName = strFileName & ".xls"
30        Call Macro2(strQ)
40        If Not fso.FolderExists(App.path & "\ExpImp") Then
50            fso.CreateFolder (App.path & "\ExpImp")
60        End If
70        If fso.FileExists(App.path & "\ExpImp\" & XlsFileName) = True Then
80            fso.DeleteFile (App.path & "\ExpImp\" & XlsFileName)
90        End If
100       Application.ActiveWorkbook.SaveAs (App.path & "\ExpImp\" & XlsFileName)
110       wkbnew.Close
120       Set wkbnew = Nothing
130       ExportIntoFile = App.path & "\ExpImp\" & XlsFileName
140   Exit Function
ERR_P:
150       MsgBox "Error in ExportIntoFile" & Err.Description & "Line : " & Erl, vbCritical
End Function

Sub Macro2(strquery As String)
On Error GoTo Err
Dim strTmp1 As String
Dim strTmp2 As String
Dim blnWithHardCode As Boolean

blnWithHardCode = True
strCurrentUserType = ADMIN
'' User Name
strTmp1 = RegGetString(HKEY_LOCAL_MACHINE, "SoftWare\ODBC\ODBC.INI\VisualStarDSN", "UID")
'' Pasword
strTmp2 = RegGetString(HKEY_LOCAL_MACHINE, "SoftWare\ODBC\ODBC.INI\VisualStarDSN", "PWD")
If strTmp1 <> "ADMIN" And InVar.strSer = 2 And blnWithHardCode = True Then
    strTmp1 = "ADMIN"
End If
If strTmp2 <> "PRINT" And InVar.strSer = 2 And blnWithHardCode = True Then
    strTmp2 = "PRINT"
End If
If InVar.strSer = 1 Then
    strTmp1 = DecryptSCR(strTmp1)
    strTmp2 = DecryptSCR(strTmp2)
End If
    Set wkbnew = Excel.Workbooks
    wkbnew.Add
    Application.Visible = False
    With ActiveSheet.QueryTables.Add(Connection:= _
        "ODBC;DSN=VisualStarDSN;Description=VisualStarDSN;UID=" & _
        strTmp1 & ";PWD=" & strTmp2 & ";DATABASE=" & Dbname & _
        ";Trusted_Connection" _
        , Destination:=Range("A1"), sql:=strquery)
        .name = "Query1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .HasAutoFormat = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TablesOnlyFromHTML = True
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
        .SavePassword = True
        .SaveData = True
    End With
Exit Sub
Err:
    Call ShowError("Error in Macro2 and Line Number : " & Erl)
End Sub

  
Public Function Export(strquery As String) As Boolean
'Start a new workbook in Excel
    Dim oApp As New Excel.Application
    Dim oBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    Dim strFileName As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim iNumCols As Integer
    rs.Open strquery, ConMain, adOpenDynamic, adLockOptimistic
    Set oBook = oApp.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
'    Dim Flt As String
'    Flt = ExpExtension(oBook.FileFormat)
    'If Flt = "FALSE" Then
    'Exit Function
    'End If
    'Add the field names in row 1
    iNumCols = rs.Fields.Count
    For i = 1 To iNumCols
        oSheet.Cells(1, i).Value = rs.Fields(i - 1).name
    Next
    'Add the data starting at cell A2
    oSheet.Range("A2").CopyFromRecordset rs
    'Format the header row as bold and autofit the columns
    With oSheet.Range("a1").Resize(1, iNumCols)
        .Font.Bold = True
        .EntireColumn.AutoFit
    End With
    oApp.Visible = True
    oApp.UserControl = True
    'Close and Recordset
    rs.Close
    Export = True
End Function

Public Function CatWiseMonPro(tmpSELCRIT As String, strDeptTmp As String) As String
    Dim strTempforCF As String
    Select Case UCase(Trim(strDeptTmp))
    Case "", "ALL"
         strTempforCF = "select Empcode,name from empmst order by Empcode"
      
    Case Else
        ''Original-> strTempforCF = "select Empcode,name from empmst Where " & SELCRIT & "=" & _
        intDeptTmp & " order by Empcode"                               'Empcode,name
        If strCurrentUserType = HOD Then
            strTempforCF = "select Empcode,name from empmst " & strCurrData & " and Empmst." & tmpSELCRIT & _
            " = " & strDeptTmp & " order by Empcode"
        Else
         strTempforCF = "select Empcode,name from empmst Where empmst." & tmpSELCRIT & _
                        " = " & strDeptTmp & " order by Empcode"
      
        End If
    End Select
    CatWiseMonPro = strTempforCF
End Function
Public Function ExpExtension(ByVal fFmt As Variant) As String   'added by  21-11
Select Case fFmt
    Case xlSYLK: ExpExtension = "*.slk|*.slk"
    Case xlWKS:  ExpExtension = "*.wks|*.wks"
    Case xlWK1, xlWK1ALL, xlWK1FMT: ExpExtension = "*.wk1|*.wk1"
    Case xlCSV, xlCSVMac, xlCSVMSDOS, xlCSVWindows: ExpExtension = "*.csv|*.csv"
    Case xlDBF2, xlDBF3, xlDBF4:   ExpExtension = "*.dbf|*.dbf"
    Case xlWorkbookNormal, xlExcel2FarEast, xlExcel3, xlExcel4, xlExcel4Workbook, xlExcel5, xlExcel7, xlExcel9795: ExpExtension = "*.xls|*.xls"
    Case xlHtml: ExpExtension = "*.htm|*.htm"
    Case xlTextMac, xlTextWindows, xlTextMSDOS, xlUnicodeText, xlCurrentPlatformText:  ExpExtension = "*.txt|*.txt"
    Case xlTextPrinter:  ExpExtension = "*.prn|*.prn"
    Case 50: ExpExtension = "*.xlsb|*.xlsb"
    Case 51: ExpExtension = "*.xlsx|*.xlsx"
    Case 52: ExpExtension = "*.xlsm|*.xlsm"
    Case 56: ExpExtension = "*.xls|*.xls"
    Case Else:  MsgBox "Do not find exact Excel File Format:: " & fFmt: ExpExtension = "FALSE"
End Select
End Function
