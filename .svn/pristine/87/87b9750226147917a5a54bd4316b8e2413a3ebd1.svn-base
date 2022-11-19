Attribute VB_Name = "mdlRep"
Option Explicit
'' Intrinsic Variables
Public strLvloc As String
Public bytRepMode As Byte
Public bytAction As Byte
Public DateStr As String
Public strCName As String
'' Yearly Report Month Name String
Public YearStr As String
''Name of the Report Selected
Public repname As Object
Public MailAdd As String
Public strRepName As String
Public RsName As Object
Public bytPoLa As Byte
'' Reports Total Leave,Present Absent,Week Off,Total Records
Public TotLeave As Single, Totpresent As Single, totAbsent As Single
Public TotWkOff As Single, TotRec As Single, TotLate As Single, TotOT As Single
Public sngLH As Single, sngEH As Single, sngOT As Single
'transaction Filename
Public strMon_Trn As String
Public strMon_Trn1 As String
Public strFName As String
Public strMon_Trn2 As String
Public strRepFile As String
Public strRepMfile As String
Public sqlStr As String    '' Used for grouping
Public headGrp As String   '' Used for sub grouping
Public StrGroup1 As String '' main group caption
Public StrGroup2 As String '' sub Group Caption
Public strSql As String    '' Used for creating query according to user selection
''Public Const rpTables = "empmst,catdesc,deptdesc,groupmst,company"
Public rpTables As String
Public strDate As String
'' Used for No Of days continues absent  (can be removed if contabs goes to periodic)
Public bytNoDay As Byte
''For Summary Report
Public intTotOnOt As Integer, sngTotOTHrs As Single
Public TempFile As String  ' 18-11
Public rpgroup1 As String   ' 18-04-09
Public strrepfile1 As String
Dim oSectionGH As CRAXDRT.Section
Dim ColArr(31) As Integer
Public cutFile As String
Public cutFile1 As String
Public strFromDt
Public strToDt
Public STRECODE As String
Public strECode1 As String
Public ReportType As String

Public Function Spaces(lenCnt)
If bytPrint = 1 Or bytPrint = 2 Or bytPrint = 3 Or (bytPrint = 4 And bytRepMode = 3) Then
        Select Case lenCnt
        Case 0: Spaces = Space(6)        '' 12
        Case 1: Spaces = Space(5)        '' 11
        Case 2: Spaces = Space(4)        '' 9
        Case 3: Spaces = Space(3)        '' 7
        Case 4: Spaces = Space(2)        '' 5
        Case 5: Spaces = Space(1)        '' 3
        Case Else: Spaces = Space(6)       '' 12
       End Select
ElseIf bytPrint = 0 Then
        Select Case lenCnt
                '' The space string here is not formed by the space key but instead with th keystroke Alt +0160
                '' Use numlock keys
                Case 0: Spaces = "      "                 '' 6Space(6)        12
                Case 1: Spaces = "     "                  '' Space(5)         11
                Case 2: Spaces = "    "                   '' Space(4)          9
                Case 3: Spaces = "   "                    '' Space(3)          7
                Case 4: Spaces = "  "                     '' Space(2)          5
                Case 5: Spaces = " "                      '' Space(1)          3
                Case Else: Spaces = "      "                '' 6Space(6)        12
        End Select
End If
End Function
': changed for monthly Reports setting the global parameter
Public Sub GetReportFile(ByVal strStruc As String, Optional bytSL As Byte)
    If bytSL <> 1 Then
    strRepFile = strStruc & intUserNum
    Else
    strRepMfile = strStruc & intUserNum
    End If
End Sub

Public Sub ChkRepFile()              '' Check for the Existence of Temporary Report File
On Error GoTo ERR_P
 

 If FindTable(strRepFile) Then ConMain.Execute "Drop table " & strRepFile
 
 If FindTable(strRepMfile) Then ConMain.Execute "Drop table " & strRepMfile
 Exit Sub
 
ERR_P:
    TruncateTable (strRepFile)
Resume Next
'    ShowError ("Check Reports File :: Reports")
End Sub

Public Sub CreRepFile(strStruc As String, Optional bytSL As Byte)   '' Create Temporary Report File
On Error GoTo ERR_P

'conmain.Execute "Select * into " & strRepFile & " from " & strStruc
If bytSL <> 1 Then
   Call ChkRepFile
 Call CreateTableIntoAs("*", strStruc, strRepFile)
     RptDel = 1
 Else
  If RptDel <> 1 Then
    Call ChkRepFile
  End If
 Call CreateTableIntoAs("*", strStruc, strRepMfile)
 RptDel = 0
 End If
If bytBackEnd = 2 Then Sleep (1000)

Exit Sub
ERR_P:
    ShowError ("Create Reports File :: Reports")
End Sub

Public Function dlyCreateFiles() As Boolean
On Error GoTo ERR_P
dlyCreateFiles = False                          '' CREATES TEMP FILES TO DUMP DATA INTO
Select Case typOptIdx.bytDly                   '' WHICH WILL BE USED BY DLYEMPSTR3() TO
    Case 0, 1, 3, 4, 5, 6, 11
        dlyCreateFiles = True
    Case 2, 8 ''conti abs,Entries
        Call GetReportFile("DPerf")
        Call CreRepFile("DPerf")
        dlyCreateFiles = True
    Case 7, 13     ''authorized / Unauthorized OT
        Call GetReportFile("Dtrn")
        Call ChkRepFile
        Call CreateTableIntoAs("*", strMon_Trn, strRepFile, " where " & strMon_Trn & "." & _
        strKDate & " = " & strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & " and " & _
        strMon_Trn & ".OTConf = '" & IIf(typOptIdx.bytDly = 7, "Y", "N") & "' ")
        Call SetZeroesToNull
        ConMain.Execute "update " & strRepFile & " set latehrs=0 where latehrs<0"
        ConMain.Execute "update " & strRepFile & " set earlhrs=0 where earlhrs<0"
        If bytBackEnd = 2 Then Sleep (1000)
        dlyCreateFiles = True
    Case 9  ''Shift Schedule
        strMon_Trn = Left(MonthName(Month(DateCompDate(typRep.strDlyDate))), 3) & _
            Right(Year(DateCompDate(typRep.strDlyDate)), 2) & "shf"
        If Not FindTable(strMon_Trn) Then
            MsgBox NewCaptionTxt("M7001", adrsMod) & MonthName(Month(DateCompDate(typRep.strDlyDate))) & _
            NewCaptionTxt("00055", adrsMod), vbInformation
            Exit Function
        End If
        dlyCreateFiles = True
    Case 10
        Call GetReportFile("DPrAb")
        Call CreRepFile("DPrAb")
        dlyCreateFiles = True
    Case 12 ''Summary Report
        Call GetReportFile("DSumC")
        Call CreRepFile("DSumC")
        dlyCreateFiles = True
        dlyCreateFiles = True
      Case Else
        MsgBox NewCaptionTxt("M7002", adrsMod), vbExclamation
        dlyCreateFiles = False
End Select
Exit Function
ERR_P:
    ShowError ("Daily Create Files :: Reports")
    ''Resume Next
End Function

Private Sub SetZeroesToNull()
On Error GoTo ERR_P
With ConMain
    .Execute ("update " & strRepFile & " set arrtim=NULL where arrtim <=0")
    .Execute ("update " & strRepFile & " set deptim=NULL where deptim <=0")
    .Execute ("update " & strRepFile & " set Latehrs=NULL where Latehrs<=0")
    .Execute ("update " & strRepFile & " set EarlHrs=NULL where EarlHrs <=0")
    .Execute ("update " & strRepFile & " set wrkhrs=NULL where wrkhrs <=0")
    .Execute ("update " & strRepFile & " set Ovtim=NULL where Ovtim <=0")
    .Execute ("update " & strRepFile & " set actrt_o = NULL where actrt_o <=0")
    .Execute ("update " & strRepFile & " set actrt_i = NULL where actrt_i <=0")
End With
Exit Sub
ERR_P:
    ShowError ("SetZeroesToNull :: mdlRep")
    ''Resume Next
End Sub
Public Function dlyTotalCalc() As Boolean
On Error GoTo RepErr
dlyTotalCalc = True
Dim strTotalCalc As String
strTotalCalc = "SELECT " & strRepFile & ".Empcode," & strRepFile & ".arrtim, " & _
strRepFile & ".latehrs," & strRepFile & ".presabs FROM " & strRepFile & "," & _
rpTables & " WHERE " & strRepFile & ".Empcode = empmst.Empcode AND " & strRepFile & _
"." & strKDate & "=" & strDTEnc & DateCompStr(typRep.strDlyDate) & strDTEnc & " " & strSql

Select Case typOptIdx.bytDly
    Case 0 '' Physical Arrival
        strTotalCalc = strTotalCalc & " AND " & strRepFile & ".arrtim>0 "
        If Not dlyTotal(0, strTotalCalc) Then
            dlyTotalCalc = False
            Exit Function
        End If
    Case 3 '' Late Arrival
        strTotalCalc = strTotalCalc & " AND " & strRepFile & ".latehrs>0 "
        If Not dlyTotal(1, strTotalCalc) Then
            dlyTotalCalc = False
            Exit Function
        End If
    Case 4 '' Early Departure
        strTotalCalc = Replace(strTotalCalc, "latehrs", "earlhrs")
        strTotalCalc = strTotalCalc & " AND " & strRepFile & ".earlhrs>0 "
        If Not dlyTotal(2, strTotalCalc) Then
            dlyTotalCalc = False
            Exit Function
        End If
    Case 6 '' Irregular
        strTotalCalc = strTotalCalc & " AND " & strRepFile & _
            ".entry < " & strRepFile & ".entreq AND " & strRepFile & ".chq ='*' "
        If Not dlyTotal(3, strTotalCalc) Then
            dlyTotalCalc = False
            Exit Function
        End If
    Case 7, 13 ''authorized / Unauthorized OT
        strTotalCalc = Replace(strTotalCalc, "latehrs", "ovtim")
        strTotalCalc = strTotalCalc & " AND " & strRepFile & ".ovtim>0 "
        If Not dlyTotal(4, strTotalCalc) Then
            dlyTotalCalc = False
            Exit Function
        End If
    Case 1 '' Absent Report
        If Not dlyTotal(5, strTotalCalc) Then
            dlyTotalCalc = False
            Exit Function
        End If
    Case 5 ''Performance
        If Not dlyTotal(5, strTotalCalc) Then
            dlyTotalCalc = False
            Exit Function
        End If
End Select
Exit Function
RepErr:
    ShowError ("Daily Total Calculations :: Reports")
    dlyTotalCalc = False
End Function

Public Function dlyTotal(ByVal bytFlag As Byte, ByVal strTotal As String) As Boolean
On Error GoTo RepErr
dlyTotal = True
'' Total Records
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strTotal, ConMain, adOpenStatic
TotRec = 0
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    TotRec = adrsTemp.RecordCount
End If
Select Case bytFlag
    Case 1 '' Late
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open strTotal, ConMain, adOpenStatic
        sngLH = 0
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not adrsTemp.EOF
                sngLH = TimAdd(sngLH, adrsTemp!latehrs)
                adrsTemp.MoveNext
            Loop
        End If
    Case 2 '' Early
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open strTotal
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            sngEH = 0
            Do While Not adrsTemp.EOF
                sngEH = TimAdd(sngEH, adrsTemp!earlhrs)
                adrsTemp.MoveNext
            Loop
        End If
    Case 0, 3, 5 '' physical Arrival,Irreg,Absent
        Totpresent = 0: totAbsent = 0: TotWkOff = 0: TotLeave = 0
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open strTotal
        If Not (adrsTemp.EOF And adrsTemp.BOF) Then
            Do While Not adrsTemp.EOF
                Select Case Left(adrsTemp!presabs, 2)
                    Case pVStar.PrsCode
                        Totpresent = Totpresent + 0.5
                    Case pVStar.AbsCode
                        totAbsent = totAbsent + 0.5
                    Case pVStar.WosCode
                        TotWkOff = TotWkOff + 0.5
                    Case pVStar.HlsCode '' Never Remove this case
                    Case Else '' Leave
                        TotLeave = TotLeave + 0.5
                End Select
                Select Case Right(adrsTemp!presabs, 2)
                    Case pVStar.PrsCode
                        Totpresent = Totpresent + 0.5
                    Case pVStar.AbsCode
                        totAbsent = totAbsent + 0.5
                    Case pVStar.WosCode
                        TotWkOff = TotWkOff + 0.5
                    Case pVStar.HlsCode '' Never remove this case
                    Case Else '' Leave
                        TotLeave = TotLeave + 0.5
                End Select
                adrsTemp.MoveNext
            Loop
        End If
    Case 4 'ot
       If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open strTotal
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            sngOT = 0
            Do While Not adrsTemp.EOF
                sngOT = TimAdd(sngOT, adrsTemp!ovtim)
                adrsTemp.MoveNext
            Loop
        End If
End Select
If adrsTemp.State = 1 Then adrsTemp.Close
Exit Function
RepErr:
    ShowError ("Daily Total :: Reports")
    dlyTotal = False
End Function

Public Function DlyEntries() As Boolean
On Error GoTo RepErr
DlyEntries = True
'' String of punches is formed from the dlydata and the record is put in DPerf table
Dim strP_Str As String, STRECODE As String
strP_Str = "": STRECODE = ""

If adrsTemp.State = 1 Then adrsTemp.Close

'    adrsTemp.Open "select  format(strf1,'hh.mm') AS t_punch ,EMPMST.Empcode   from  TBLDATA," & _
'    rpTables & " where empmst.Empcode=TBLDATA.STRCODE " & strSql & _
'     " order by EMPMST.Empcode,   Format(strf1,'hh.nn') ", conmain, adOpenStatic

    adrsTemp.Open "select  format(strf1,'hh.mm') AS t_punch ,EMPMST.Empcode   from  TBLDATA," & _
    rpTables & " where VAL(empmst.CARD)=VAL(LEFT(TBLDATA.STRCODE, " & pVStar.CardSize & ")) " & strSql & _
     " order by EMPMST.Empcode,   Format(strf1,'hh.nn') ", ConMain, adOpenStatic


If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    adrsTemp.MoveFirst
    Do While Not adrsTemp.EOF
        STRECODE = adrsTemp!Empcode
        strP_Str = ""
        Do  '' Same Empcode
            strP_Str = strP_Str & IIf(Not IsNull(adrsTemp!t_punch) And _
            adrsTemp!t_punch > 0, Format(adrsTemp!t_punch, "0.00") & _
            Spaces(4), Spaces(4))
            adrsTemp.MoveNext
            If adrsTemp.EOF Then Exit Do
        Loop Until STRECODE <> adrsTemp!Empcode And Not adrsTemp.EOF
        
        If Len(strP_Str) > 255 Then strP_Str = Left(strP_Str, 250)
        If Trim(strP_Str) <> "" Then
            ConMain.Execute "insert into " & strRepFile & _
            "(Empcode,punches) values" & "(" & "'" & STRECODE & "'" & ", " & "'" & _
            strP_Str & "'" & ")"
        Else
            ConMain.Execute "insert into " & strRepFile & _
            "(Empcode,punches) values" & "(" & "'" & STRECODE & "'" & ", " & "'" & _
            strP_Str & "'" & ")"
        End If
    Loop
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    DlyEntries = False
End If
If adrsTemp.State = 1 Then adrsTemp.Close
Exit Function
RepErr:
    ShowError ("Daily Entries :: Reports")
    DlyEntries = False
End Function

Public Function MonthEntries() As Boolean

MonthEntries = True
'' String of punches is formed from the dlydata and the record is put in DPerf table
Dim strP_Str As String, STRECODE As String
strP_Str = "": STRECODE = ""
Dim PunchDT As String
Dim EntriExist As Boolean

Dim rsEmpcode As New ADODB.Recordset
Dim rsEntries As New ADODB.Recordset
If rsEmpcode.State = 1 Then rsEmpcode.Close
rsEmpcode.Open "Select distinct StrCode from tblData", ConMain, adOpenStatic
TruncateTable ("ImportTbl")
Do While Not rsEmpcode.EOF
    Dim PreviousPunch As Single
    EntriExist = False
    If rsEntries.State = 1 Then rsEntries.Close
    rsEntries.Open "SELECT format(strf1,'dd') AS DT, format(strf1,'hh.mm') AS t_punch FROM TBLDATA Where strcode = '" & rsEmpcode.Fields("strcode") & "' ORDER BY  format(strf1,'dd'), Format(strf1,'hh.nn')", ConMain, adOpenStatic
    
    rsEntries.MoveFirst
    Do While Not rsEntries.EOF
    strP_Str = ""
    PunchDT = Val(rsEntries!dt)
      
        Do  '' Same Empcode
            strP_Str = strP_Str & IIf(Not IsNull(rsEntries!t_punch) And _
            rsEntries!t_punch > 0, Format(rsEntries!t_punch, "0.00") & _
            Spaces(5), Spaces(5))
            
            PreviousPunch = rsEntries!t_punch
            rsEntries.MoveNext
            If rsEntries.EOF Then Exit Do
            If Val(rsEntries!t_punch) = Val(PreviousPunch) Then rsEntries.MoveNext
            If rsEntries.EOF Then Exit Do
        Loop Until Val(PunchDT) <> Val(rsEntries!dt) And Not rsEntries.EOF
        If Trim(strP_Str) <> "" Then
            If EntriExist = False Then
                ConMain.Execute "INSERT INTO ImportTbl ( empcode, D" & PunchDT & " ) VALUES ('" & rsEmpcode.Fields("strcode") & "', '" & Left(strP_Str, 250) & "')"
                EntriExist = True
            Else
                    ConMain.Execute "Update ImportTbl set  D" & PunchDT & " = '" & Left(strP_Str, 250) & "' Where empcode = '" & rsEmpcode!strCode & "'"
            End If
        End If
    
   Loop
    rsEmpcode.MoveNext
Loop

End Function


Public Function MonEntries() As Boolean 'Added by  11-07 monthly entries
On Error GoTo RepErr
Dim rsTemp As New ADODB.Recordset
MonEntries = True
Dim strP_Str As String, STRECODE As String
Dim Mnth As String, Yr As String
Dim dte As Date
strP_Str = "": STRECODE = ""
Mnth = MonthNumber(typRep.strMonMth)
Yr = typRep.strMonYear

If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select distinct dailypro.Empcode,dailypro.dte,dailypro.t_punch from dailypro," & _
rpTables & " where empmst.Empcode=dailypro.Empcode and  month(dailypro.dte) =" & Mnth & " and year(dailypro.dte)=" & Yr & " " & strSql & " order by dailypro.Empcode," & _
"dailypro.dte,dailypro.t_punch ", ConMain, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    adrsTemp.MoveFirst
    Do While Not adrsTemp.EOF
        STRECODE = adrsTemp!Empcode
        dte = adrsTemp!dte
        strP_Str = ""
        Do  '' Same Empcode
            strP_Str = strP_Str & IIf(Not IsNull(adrsTemp!t_punch) And _
            adrsTemp!t_punch > 0, Format(adrsTemp!t_punch, "0.00") & _
            Spaces(Len(Format(adrsTemp!t_punch, "0.00"))), "")
            adrsTemp.MoveNext
            If adrsTemp.EOF Then Exit Do
        Loop Until (STRECODE <> adrsTemp!Empcode Or dte <> adrsTemp!dte) And Not adrsTemp.EOF
        If Trim(strP_Str) <> "" Then
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open "select empcode from " & strRepFile & " where empcode='" & STRECODE & "'", ConMain, adOpenStatic
            If (rsTemp.EOF And rsTemp.BOF) Then
                ConMain.Execute "insert into " & strRepFile & _
                "(empcode,P" & Day(dte) & ") values" & "(" & "'" & STRECODE & "'" & ",'" & strP_Str & "')"
            Else
                ConMain.Execute "Update " & strRepFile & " set P" & Day(dte) & "='" & strP_Str & "' where empcode='" & STRECODE & "'"
            End If
        End If
    Loop
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    MonEntries = False
End If
If adrsTemp.State = 1 Then adrsTemp.Close
Exit Function
RepErr:
    ShowError ("Daily Entries :: Reports")
    MonEntries = False
    'Resume Next
End Function

Public Function dlyContAbs(ByVal strFrDate As String, ByVal strLaDate As String) As Boolean
On Error GoTo RepErr
dlyContAbs = True
Dim bytTempCnt As Byte
Dim dte As String, strFileName1 As String, strFileName2 As String
Dim strP_Str As String, strTempdt As String, STRECODE As String

strFileName1 = MakeName(MonthName(Month(DateCompDate(strFrDate))), Year(DateCompDate(strFrDate)), "trn")
strFileName2 = MakeName(MonthName(Month(DateCompDate(strLaDate))), Year(DateCompDate(strLaDate)), "trn")
bytTempCnt = 0
Dim strGP As String
    If strFileName1 = strFileName2 Then
        strGP = "select " & strFileName1 & ".Empcode,presabs," & strKDate & " from " & strFileName1 & _
        "," & rpTables & " where " & strFileName1 & ".Empcode = empmst.Empcode and " & _
        "" & strKDate & " >=" & strDTEnc & DateCompStr(strFrDate) & strDTEnc & " and " & strKDate & "<=" & _
        strDTEnc & DateCompStr(strLaDate) & strDTEnc & " And ( presabs <> '" & _
        pVStar.PrsCode & pVStar.PrsCode & "' ) " & strSql
    Else
        strGP = "select  " & strFileName1 & ".Empcode,presabs," & strKDate & " from " & strFileName1 & _
        "," & rpTables & " where " & strFileName1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & _
        strDTEnc & DateCompStr(strFrDate) & strDTEnc & " And presabs <> '" & pVStar.PrsCode & pVStar.PrsCode & _
        "' " & strSql & _
        " union  select " & strFileName2 & ".Empcode,presabs," & strKDate & " from  " & strFileName2 & _
        "," & rpTables & " where " & strFileName2 & ".Empcode = empmst.Empcode and " & strKDate & "<=" & _
        strDTEnc & DateCompStr(strLaDate) & strDTEnc & " And  presabs <>'" & pVStar.PrsCode & _
        pVStar.PrsCode & "' " & strSql
    End If
    Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strFileName1 & ".Empcode," & strFileName1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
    End Select

    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open strGP, ConMain, adOpenStatic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
        adrsTemp.MoveFirst
        dte = CDate(strFrDate)
        DateStr = ""
        strTempdt = ""
        If Day(strFrDate) >= Day(strLaDate) Then
            strTempdt = strFrDate
            Do While CDate(strTempdt) <= CDate(strLaDate)
                DateStr = DateStr & Day(CDate(strTempdt)) & Spaces(Len(Trim(str(Day(CDate(strTempdt))))))
                strTempdt = CStr(CDate(strTempdt) + 1)
            Loop
        Else
            i = Day(strFrDate)
            For i = i To Day(strLaDate)
                DateStr = DateStr & i & Spaces(Len(Trim(CStr(i))))
            Next i
        End If
         Do While Not (adrsTemp.EOF) And dte <= CDate(strLaDate)
            STRECODE = adrsTemp!Empcode
            dte = CDate(strFrDate)
            strP_Str$ = ""
            Do While dte >= CDate(strFrDate) And dte <= CDate(strLaDate) And Not (adrsTemp.EOF)
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dte Then
                        If adrsTemp!presabs <> pVStar.AbsCode & pVStar.AbsCode And _
                        adrsTemp!presabs <> pVStar.WosCode & pVStar.WosCode And _
                        adrsTemp!presabs <> pVStar.HlsCode & pVStar.HlsCode Then
                            strP_Str = ""
                            Do
                                adrsTemp.MoveNext
                                If adrsTemp.EOF Then Exit Do
                            Loop Until adrsTemp!Empcode <> STRECODE
                            Exit Do
                        Else
                            strP_Str = strP_Str & IIf(Not IsNull(adrsTemp!presabs), _
                            adrsTemp!presabs, "") & Spaces(Len(adrsTemp!presabs))
                            bytTempCnt = bytTempCnt + 1
                        End If
                    ElseIf adrsTemp!Date <> dte Then
                        strP_Str = strP_Str & Spaces(0)
                    End If
                Else
                    Exit Do
                End If
                If dte = adrsTemp!Date Then
                    adrsTemp.MoveNext
                    'dte = CDate(strFrDate)
                    dte = DateAdd("d", 1, dte)
                Else
                    dte = DateAdd("d", 1, dte)
                End If
            Loop
        If strP_Str <> "" And bytTempCnt = bytNoDay Then
            ConMain.Execute "insert into " & strRepFile & "(Empcode," & _
            "presabsstr)  values('" & STRECODE & "','" & strP_Str & "')"
        End If
        bytTempCnt = 0
        dte = CDate(strFrDate)
    Loop
'DateStr = "" ' NEVER UNCOMMENT OR USE THIS STATEMENT.THIS VALUE IS USED BY RELATED DSR
Else '' If No Records found
    Call SetMSF1Cap(10)
    dlyContAbs = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If
If adrsTemp.State = 1 Then adrsTemp.Close
Exit Function
RepErr:
    ShowError ("Daily Continuos Absent :: Reports")
    dlyContAbs = False
End Function

Public Function DlyManpower() As Boolean
On Error GoTo RepErr
DlyManpower = True
Dim bytCnt As Integer
Dim strTmp As String
Select Case sqlStr          ''for serial no.this will give serial
    Case "Empcode"          '' no according to the grouping.
    Case "deptdescdept": strTmp = "deptdesc.dept,"
    Case "catdesccat": strTmp = "catdesc.cat,"
    Case "groupmst": strTmp = "groupmst.grupdesc,"
    Case "deptdescdept','deptdescdesc": strTmp = "deptdesc.dept,catdesc.cat,"
End Select
    strMon_Trn = MakeName(MonthName(Month(DateCompDate(typRep.strDlyDate))), Year(DateCompDate(typRep.strDlyDate)), "trn")
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select " & strMon_Trn & ".Empcode,presabs from " & strMon_Trn & _
    "," & rpTables & " where " & strKDate & "=" & strDTEnc & DateCompStr(typRep.strDlyDate) & _
    strDTEnc & " and empmst.Empcode=" & strMon_Trn & ".Empcode " & strSql & _
    " Order by " & strTmp & " empmst.Empcode ", ConMain, adOpenStatic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
        totAbsent = 0: Totpresent = 0: TotWkOff = 0
        Do While Not adrsTemp.EOF
            bytCnt = bytCnt + 1
            Select Case adrsTemp!presabs
                Case pVStar.AbsCode & pVStar.AbsCode:
                    ConMain.Execute "insert into " & strRepFile & "(srno,Empcode,absent,absentT) " & _
                    " values (" & bytCnt & ",'" & adrsTemp!Empcode & "','" & adrsTemp!presabs & "',1)"
                    totAbsent = totAbsent + 1
                Case pVStar.AbsCode & pVStar.PrsCode, pVStar.PrsCode & pVStar.AbsCode
                    ConMain.Execute "insert into " & strRepFile & "(srno,Empcode,absent,present," & _
                    " absentT,presentT) values (" & bytCnt & ",'" & adrsTemp!Empcode & "','" & adrsTemp!presabs & _
                    "','" & adrsTemp!presabs & "',0.5,0.5)"
                    totAbsent = totAbsent + 0.5
                    Totpresent = Totpresent + 0.5
                Case pVStar.PrsCode & pVStar.PrsCode:
                    ConMain.Execute "insert into " & strRepFile & "(srno,Empcode,present,presentT) " & _
                    " values (" & bytCnt & ",'" & adrsTemp!Empcode & "','" & adrsTemp!presabs & "',1)"
                    Totpresent = Totpresent + 1
                Case pVStar.WosCode & pVStar.WosCode:
                    ConMain.Execute "insert into " & strRepFile & "(srno,Empcode,offs,offsT) values (" & _
                    bytCnt & ",'" & adrsTemp!Empcode & "','" & adrsTemp!presabs & "',1)"
                    TotWkOff = TotWkOff + 1
                Case Else
                    bytCnt = bytCnt - 1
            End Select
            adrsTemp.MoveNext
        Loop
    Else '' If No Records found
        Call SetMSF1Cap(10)
        DlyManpower = False
        MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    End If
    If adrsTemp.State = 1 Then adrsTemp.Close
Exit Function
RepErr:
        ShowError ("Daily Manpower :: Reports")
        DlyManpower = False
        'Resume Next
End Function
'' for Summery Report
Public Function Fuc_NewSummary() As Boolean
Dim strFileName As String
Dim sngpp As Single, sngAA As Single, sngWO As Single, sngHL As Single
Dim snglv As Single, sngOD As Single, sngTot As Single
Dim sngPPP As Single, sngAAP As Single, sngWOP As Single, sngHLP As Single
Dim sngLVP As Single, sngODP As Single, sngTotP As Single, sngOT As Single, sngOtHrs As Single
Dim bytCnt As Single, strDept As String, intStrength As Integer
Dim sngGTot As String, Msum As String, sngshift As String
Dim tmp  As String, MStrengthTot As String
Dim MOrderby As String
Dim Query, mf, mv As String
Dim Mcnt, cnt As Byte
Dim Mcnt1, Mcnt2() As String
Dim mt As Variant

If typOptIdx.bytDly = 12 Or typOptIdx.bytDly = 32 Then
    strMon_Trn = Left(MonthName(Month(DateCompDate(typRep.strDlyDate))), 3) & _
            Right(Year(DateCompDate(typRep.strDlyDate)), 2) & "trn"
Else
    strMon_Trn = Left(MonthName(Month(DateCompDate(typRep.strPeriFr))), 3) & _
            Right(Year(DateCompDate(typRep.strPeriFr)), 2) & "trn"
End If

For i = 1 To 7
    If frmReports.lblGrp(i).Caption <> "" Then
        If frmReports.optGrp(i).Caption = "Employee" Then
            If (MOrderby) = "" Then
                MOrderby = "empmst.Empcode"
            Else
              MOrderby = MOrderby & "empmst.Empcode"
            End If
        ElseIf frmReports.optGrp(i).Caption = "Category" Then
            If (MOrderby) = "" Then
                MOrderby = "catdesc.cat"
            Else
                MOrderby = MOrderby & " , catdesc.cat"
            End If
        ElseIf frmReports.optGrp(i).Caption = "Department" Then
            If (MOrderby) = "" Then
                MOrderby = " deptdesc.dept"
            Else
               MOrderby = MOrderby & " , deptdesc.dept"
            End If
        ElseIf frmReports.optGrp(i).Caption = "Location" Then
            If (MOrderby) = "" Then
                MOrderby = " Location.Location"
            Else
                MOrderby = MOrderby & " ,Location.Location"
            End If
        ElseIf frmReports.optGrp(i).Caption = "Division" Then
            If (MOrderby) = "" Then
                MOrderby = " Division.Div"
            Else
                MOrderby = MOrderby & " ,Division.Div"
            End If
        ElseIf frmReports.optGrp(i).Caption = "Group" Then
            If (MOrderby) = "" Then
                MOrderby = " groupmst." & strKGroup & ""
            Else
                MOrderby = MOrderby & " , groupmst." & strKGroup & ""
            End If
        ElseIf frmReports.optGrp(i).Caption = "Company" Then
            If (MOrderby) = "" Then
                MOrderby = " Company.Company "
            Else
            'Abe  aisi galti karega to kaise chalega "Groupmst.company likha thaa idhar"
                MOrderby = MOrderby & " , Company.Company"
            End If
        
        End If
    End If
Next i

If adrsTemp.State = 1 Then adrsTemp.Close

If typOptIdx.bytDly = 12 Or typOptIdx.bytDly = 32 Then
adrsTemp.Open "Select DISTINCT Deptdesc.dept as dept ,Deptdesc.strenth as strenth,Deptdesc." & strKDesc & " as deptdescdesc  ,  " & _
    "empmst.Empcode  as Empcode , empmst.name as Name ,catdesc.cat as cat,catdesc." & strKDesc & "  as catdescdesc , " & _
    "Location.Location as Location , Location.locDesc as LocDesc , Division.Div  as Div , Division.Divdesc as Divdesc , " & _
    "groupmst." & strKGroup & " as " & strKGroup & " , groupmst.GrupDesc as Grupdesc , Company.Company as Company," & _
    " Company.cname as Cname ," & strMon_Trn & ".shift ," & strMon_Trn & ".ovtim , " & strMon_Trn & _
    ".PRESABS," & strMon_Trn & ".EmpCode, " & strMon_Trn & "." & strKDate & ", " & strMon_Trn & _
    ".od_from From " & rpTables & "," & strMon_Trn & " WHERE empmst.Empcode = " & _
    strMon_Trn & ".EMPCODE and " & strMon_Trn & "." & strKDate & "= " & strDTEnc & _
    DateCompStr(typRep.strDlyDate) & strDTEnc & " " & strSql & "  ORDER BY " & _
    strMon_Trn & "." & strKDate & "," & strMon_Trn & ".shift", ConMain, adOpenStatic
Else
adrsTemp.Open "Select DISTINCT Deptdesc.dept as dept , Deptdesc." & strKDesc & " as deptdescdesc  ,  " & _
    "empmst.Empcode  as Empcode , empmst.name as Name ,catdesc.cat as cat,catdesc." & strKDesc & "  as catdescdesc , " & _
    "Location.Location as Location , Location.locDesc as LocDesc , Division.Div  as Div , Division.Divdesc as Divdesc , " & _
    "groupmst." & strKGroup & " as " & strKGroup & " , groupmst.GrupDesc as Grupdesc , Company.Company as Company, " & _
    "Company.cname as Cname ," & strMon_Trn & ".ovtim  , " & strMon_Trn & _
    ".PRESABS," & strMon_Trn & ".EmpCode, " & strMon_Trn & "." & strKDate & ", " & strMon_Trn & _
    ".od_from From " & rpTables & "," & strMon_Trn & " WHERE empmst.Empcode = " & _
    strMon_Trn & ".EMPCODE and " & strMon_Trn & "." & strKDate & " between " & strDTEnc & _
    DateCompStr(typRep.strPeriFr) & strDTEnc & "and " & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & _
    " " & strSql & "  ORDER BY " & MOrderby & "," & strMon_Trn & "." & strKDate, ConMain, adOpenStatic
End If
ReDim Mcnt2(8) As String
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    TotLeave = 0: Totpresent = 0: totAbsent = 0
    TotWkOff = 0: TotRec = 0: TotLate = 0: TotOT = 0
    Do While Not adrsTemp.EOF
    Dim MRs_value(8) As String, MSet_Value(8) As String, Mfield_nm(8) As String, Mdesc(8) As String, Mdesc_Nm(8) As String
        Mcnt = 1
        For i = 1 To 7
            If frmReports.lblGrp(i).Caption <> "" Then
                If frmReports.optGrp(i).Caption = "Employee" Then
                    MSet_Value(Mcnt) = adrsTemp("Empcode"): Mcnt2(Mcnt) = 0
                    Mfield_nm(Mcnt) = "Empcode": Mdesc_Nm(Mcnt) = "EmpName"
                    Mdesc(Mcnt) = adrsTemp("Name"): Mcnt = Mcnt + 1
                    If (MStrengthTot) = "" Then
                        MStrengthTot = "Empcode='" & adrsTemp("Empcode") & "'"
                    Else
                        MStrengthTot = MStrengthTot & " And  Empcode='" & adrsTemp("Empcode") & "'"
                    End If
                ElseIf frmReports.optGrp(i).Caption = "Category" Then
                    MSet_Value(Mcnt) = adrsTemp("cat"): Mcnt2(Mcnt) = 1
                    Mfield_nm(Mcnt) = "cat": Mdesc_Nm(Mcnt) = "catdescdesc"
                    Mdesc(Mcnt) = adrsTemp("catdescdesc"): Mcnt = Mcnt + 1
                    If (MStrengthTot) = "" Then
                        MStrengthTot = "cat='" & adrsTemp("cat") & "'"
                    Else
                        MStrengthTot = MStrengthTot & " And cat='" & adrsTemp("cat") & "'"
                    End If
                ElseIf frmReports.optGrp(i).Caption = "Department" Then
                    MSet_Value(Mcnt) = adrsTemp("dept"): Mcnt2(Mcnt) = 2
                    Mfield_nm(Mcnt) = "dept": Mdesc_Nm(Mcnt) = "deptdescdesc"
                    Mdesc(Mcnt) = adrsTemp("deptdescdesc"): Mcnt = Mcnt + 1
                    If (MStrengthTot) = "" Then
                        MStrengthTot = "dept=" & adrsTemp("dept") & ""
                    Else
                        MStrengthTot = MStrengthTot & " And  dept=" & adrsTemp("dept") & ""
                    End If
                ElseIf frmReports.optGrp(i).Caption = "Location" Then
                    MSet_Value(Mcnt) = adrsTemp("Location"): Mcnt2(Mcnt) = 3
                    Mfield_nm(Mcnt) = "Location": Mdesc_Nm(Mcnt) = "LocDesc"
                    Mdesc(Mcnt) = adrsTemp("LocDesc"): Mcnt = Mcnt + 1
                    If (MStrengthTot) = "" Then
                        MStrengthTot = "Location=" & adrsTemp("Location")
                    Else
                        MStrengthTot = MStrengthTot & " And  Location=" & adrsTemp("Location")
                    End If
                ElseIf frmReports.optGrp(i).Caption = "Division" Then
                    MSet_Value(Mcnt) = adrsTemp("Div"): Mcnt2(Mcnt) = 4
                    Mfield_nm(Mcnt) = "div": Mdesc_Nm(Mcnt) = "Divdesc"
                    Mdesc(Mcnt) = adrsTemp("Divdesc"): Mcnt = Mcnt + 1
                    If (MStrengthTot) = "" Then
                        MStrengthTot = "div=" & adrsTemp("div")
                    Else
                        MStrengthTot = MStrengthTot & " And  div=" & adrsTemp("div")
                    End If
                ElseIf frmReports.optGrp(i).Caption = "Group" Then
                    MSet_Value(Mcnt) = adrsTemp("Group"): Mcnt2(Mcnt) = 5
                    Mfield_nm(Mcnt) = "" & strKGroup & "": Mdesc_Nm(Mcnt) = "grupdesc"
                    Mdesc(Mcnt) = adrsTemp("grupdesc"): Mcnt = Mcnt + 1
                    If (MStrengthTot) = "" Then
                        MStrengthTot = "" & strKGroup & "=" & adrsTemp!Group & ""
                    Else
                        MStrengthTot = MStrengthTot & " And  " & strKGroup & "=" & adrsTemp!Group & ""
                    End If
                ElseIf frmReports.optGrp(i).Caption = "Company" Then
                    MSet_Value(Mcnt) = adrsTemp("Company"): Mcnt2(Mcnt) = 6
                    Mfield_nm(Mcnt) = "Company": Mdesc_Nm(Mcnt) = "Cname"
                    Mdesc(Mcnt) = adrsTemp("Cname"): Mcnt = Mcnt + 1
                    If (MStrengthTot) = "" Then
                        MStrengthTot = "Company=" & adrsTemp("Company")
                    Else
                        MStrengthTot = MStrengthTot & " And  Company=" & adrsTemp("Company")
                    End If

               End If
            End If
        Next i
mt = Array(adrsTemp("Empcode"), adrsTemp("cat"), adrsTemp("dept"), adrsTemp("Location"), adrsTemp("div"), adrsTemp("group"), adrsTemp("Company"))
        Mcnt = Mcnt - 1
'        If typOptIdx.bytDly = 32 Then
'        MStrengthTot = "Shift='" & adrsTemp("shift") & "'"
'        End If
        intStrength = CgetSTR(MStrengthTot)
        Do While IIf(Mcnt = 1, MSet_Value(1) = mt(Val(Mcnt2(1))), _
                 IIf(Mcnt = 2, MSet_Value(1) = mt(Val(Mcnt2(1))) And MSet_Value(2) = mt(Val(Mcnt2(2))), _
                 IIf(Mcnt = 3, MSet_Value(1) = mt(Val(Mcnt2(1))) And MSet_Value(2) = mt(Val(Mcnt2(2))) And MSet_Value(3) = mt(Val(Mcnt2(3))), _
                 IIf(Mcnt = 4, MSet_Value(1) = mt(Val(Mcnt2(1))) And MSet_Value(2) = mt(Val(Mcnt2(2))) And MSet_Value(3) = mt(Val(Mcnt2(3))) And MSet_Value(4) = mt(Val(Mcnt2(4))), _
                 IIf(Mcnt = 5, MSet_Value(1) = mt(Val(Mcnt2(1))) And MSet_Value(2) = mt(Val(Mcnt2(2))) And MSet_Value(3) = mt(Val(Mcnt2(3))) And MSet_Value(4) = mt(Val(Mcnt2(4))) And MSet_Value(5) = mt(Val(Mcnt2(5))), _
                 IIf(Mcnt = 6, MSet_Value(1) = mt(Val(Mcnt2(1))) And MSet_Value(2) = mt(Val(Mcnt2(2))) And MSet_Value(3) = mt(Val(Mcnt2(3))) And MSet_Value(4) = mt(Val(Mcnt2(4))) And MSet_Value(5) = mt(Val(Mcnt2(5))) And MSet_Value(6) = mt(Val(Mcnt2(6))), _
                 IIf(Mcnt = 7, MSet_Value(1) = mt(Val(Mcnt2(1))) And MSet_Value(2) = mt(Val(Mcnt2(2))) And MSet_Value(3) = mt(Val(Mcnt2(3))) And MSet_Value(4) = mt(Val(Mcnt2(4))) And MSet_Value(5) = mt(Val(Mcnt2(5))) And MSet_Value(6) = mt(Val(Mcnt2(6))) And MSet_Value(7) = mt(Val(Mcnt2(7))), "")))))))
                 'IIf(Mcnt = 8, MSet_Value(1) = mt(Val(Mcnt2(1))) And MSet_Value(2) = mt(Val(Mcnt2(2))) And MSet_Value(3) = mt(Val(Mcnt2(3))) And MSet_Value(4) = mt(Val(Mcnt2(4))) And MSet_Value(5) = mt(Val(Mcnt2(5))) And MSet_Value(6) = mt(Val(Mcnt2(6))) And MSet_Value(7) = mt(Val(Mcnt2(7))) And MSet_Value(8) = mt(Val(Mcnt2(8))), ""))))))))
            Select Case Left(adrsTemp("presabs"), 2)
                Case pVStar.PrsCode
                    sngpp = sngpp + 0.5
                Case pVStar.AbsCode
                    sngAA = sngAA + 0.5
                Case pVStar.WosCode
                    sngWO = sngWO + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case "OD"
                    sngOD = sngOD + 0.5
                Case Else
                    snglv = snglv + 0.5
            End Select
            Select Case Right(adrsTemp("presabs"), 2)
                Case pVStar.PrsCode
                    sngpp = sngpp + 0.5
                Case pVStar.AbsCode
                    sngAA = sngAA + 0.5
                Case pVStar.WosCode
                    sngWO = sngWO + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case "OD"
                    sngOD = sngOD + 0.5
                Case Else
                    snglv = snglv + 0.5
            End Select
            sngTot = sngTot + 1
'            If adrsTemp("od_from") > 0 Then sngOD = sngOD + 1
            If adrsTemp("ovtim") > 0 Then sngOT = sngOT + 1
            If adrsTemp("ovtim") > 0 Then sngOtHrs = TimAdd(sngOtHrs, adrsTemp!ovtim)
            If typOptIdx.bytDly = 32 Then
            If adrsTemp("shift") <> "" Then sngshift = adrsTemp("shift")
            End If
            adrsTemp.MoveNext
            If adrsTemp.EOF = True Then Exit Do
        Loop
        bytCnt = bytCnt + 1
        If sngpp > 0 Then sngPPP = Format((sngpp * 100) / sngTot, "00.00")
        If sngAA > 0 Then sngAAP = Format((sngAA * 100) / sngTot, "00.00")
        If sngWO > 0 Then sngWOP = Format((sngWO * 100) / sngTot, "00.00")
        If sngHL > 0 Then sngHLP = Format((sngHL * 100) / sngTot, "00.00")
        If snglv > 0 Then sngLVP = Format((snglv * 100) / sngTot, "00.00")
        sngGTot = Round(sngPPP + sngAAP + sngWOP + sngHLP + sngLVP)
      
        ''For DSR
        Totpresent = Totpresent + sngpp     ''PP total
        totAbsent = totAbsent + sngAA       ''AA total
        TotWkOff = TotWkOff + sngWO         ''WO total
        TotOT = TotOT + sngHL               ''HL total
        TotLeave = TotLeave + snglv         ''LV total
        TotLate = TotLate + sngOD           ''OD total
        
        Query = ""
        mf = ""
        mv = ""
        Query = "Insert into " & strRepFile & "("
         
         For cnt = 1 To UBound(Mfield_nm)
            If Mfield_nm(cnt) <> "" Then
                mf = mf & IIf(cnt = 1, "", " , ") & Mfield_nm(cnt) & " ," & Mdesc_Nm(cnt)
                 If IsNumeric(MSet_Value(cnt)) Then
                     mv = mv & IIf(cnt = 1, "", " , ") & MSet_Value(cnt)
                 Else
                        If MSet_Value(cnt) = "" Then
                            mv = mv & IIf(cnt = 1, "''", " , ''")
                        Else
                            mv = mv & IIf(cnt = 1, "'", " , '") & MSet_Value(cnt) & "' "
                        End If
                 End If
                    If IsNumeric(Mdesc(cnt)) Then
                        mv = mv & " , " & Mdesc(cnt)
                    Else
                        If Mdesc(cnt) = "" Then
                            mv = mv & " , ''"
                        Else
                            mv = mv & " , '" & Mdesc(cnt) & "' "
                        End If
                    End If
                 End If
         Next
            mf = mf & ",Strength,NofE,PP,PPP,AA,AAP,WO,WOP,HL,HLP,PL,PLP,UPL,UPLP,OD,ODP,TOT,TOTP,OT,OTHRS,shift)"
            mv = mv & "," & intStrength & "," & sngTot & "," & sngpp & "," & sngPPP & "," & sngAA & "," & sngAAP & "," & sngWO & "," & sngWOP & _
                "," & sngHL & "," & sngHLP & "," & snglv & "," & sngLVP & ",0,0," & sngOD & ",0," & sngTot & "," & sngGTot & "," & sngOT & "," & sngOtHrs & ",'" & sngshift & "'"
             
            Query = Query & mf & " Values (" & mv & ")"
            ConMain.Execute Query
                                                                                                                   
          
        sngpp = 0: sngAA = 0: sngWO = 0: sngHL = 0
        snglv = 0: sngOD = 0: sngTot = 0
        sngPPP = 0: sngAAP = 0: sngWOP = 0: sngHLP = 0
        sngLVP = 0: sngODP = 0: sngTotP = 0
        sngOT = 0: sngOtHrs = 0
        MStrengthTot = ""
        
    Loop
Fuc_NewSummary = True
Else
End If
End Function

Private Function CgetSTR(ByVal PString As String) As Integer
On Error GoTo ERR_P
'MsgBox PString
If AdrsCat.State = 1 Then AdrsCat.Close
AdrsCat.Open "Select count(*) from empmst where  " & PString & " and leavdate is null ", ConMain, adOpenStatic
If Not (AdrsCat.EOF And AdrsCat.BOF) Then
    CgetSTR = IIf(IsNull(AdrsCat(0)), 0, AdrsCat(0))
End If
Exit Function
ERR_P:
ShowError ("CgetSTR :: Summary Report")
End Function

Public Function dlyPunchVari() As Boolean
On Error GoTo ERR_P
Dim sngDiff As Single
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "Select * from " & strRepFile & " order by Empcode", _
ConMain, adOpenStatic, adLockOptimistic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    Do While Not adrsTemp.EOF
        sngDiff = TimDiff(adrsTemp("DEPTIM"), adrsTemp("actrt_i"))
        adrsTemp("actrt_o") = sngDiff
        adrsTemp.Update
        adrsTemp.MoveNext
    Loop
    dlyPunchVari = True
Else
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("dlyPunchVari : mdlRep")
End Function

Private Function getSTR(ByVal intDept As Integer) As Integer
On Error GoTo ERR_P
If AdrsCat.State = 1 Then AdrsCat.Close
AdrsCat.Open "Select count(*) from empmst where dept = " & intDept, ConMain, adOpenStatic
If Not (AdrsCat.EOF And AdrsCat.BOF) Then
    getSTR = IIf(IsNull(AdrsCat(0)), 0, AdrsCat(0))
End If
Exit Function
ERR_P:
ShowError ("getSTR :: Summary Report")
End Function

Public Function monDateStr(ByVal bytDay As Byte) As String
Select Case bytAction
    Case 2, 3, 4
        Select Case bytDay
            Case 28
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28"
            Case 29
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28    29"
            Case 30
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28    29    30"
            Case 31
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28    29    30    31"
        End Select
    Case 1
        Select Case bytDay
            Case 28
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28"
            Case 29
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28    29"
            Case 30
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28    29    30"
            Case 31
                monDateStr = "1     2     3     4     5     6     7     8     9     10    11    12    13    14    15    16    17    18    19    20    21    22    23    24    25    26    27    28    29    30    31"
        End Select
End Select
End Function

Public Function FdtLdt(ByVal MonthByt As Byte, ByVal CurYear As String, Optional ByVal Ofchr As String = "F") As String
On Error GoTo ERR_P
If UCase(Ofchr) = "F" Then
    Select Case bytDateF
       Case 1      '' American
           FdtLdt = (Format(MonthByt, "00") & "/01/" & CurYear)
       Case 2      '' British
           FdtLdt = ("01/" & Format(MonthByt, "00") & "/" & CurYear)
    End Select
    Exit Function
Else
    If MonthByt = 0 Then MonthByt = 12
    If MonthByt = 13 Then MonthByt = 1
    Select Case MonthByt
        Case 1, 3, 5, 7, 8, 10, 12:
            Select Case bytDateF
                Case 1      '' American
                    FdtLdt = (Format(MonthByt, "00") & "/31/" & CurYear)
                Case 2      '' British
                    FdtLdt = ("31/" & Format(MonthByt, "00") & "/" & CurYear)
            End Select
        Case 4, 6, 9, 11:
            Select Case bytDateF
                Case 1      '' American
                    FdtLdt = (Format(MonthByt, "00") & "/30/" & CurYear)
                Case 2      '' British
                    FdtLdt = ("30/" & Format(MonthByt, "00") & "/" & CurYear)
            End Select
        Case 2:
            If CurYear Mod 4 = 0 Then
                Select Case bytDateF
                Case 1      '' American
                    FdtLdt = (Format(MonthByt, "00") & "/29/" & CurYear)
                Case 2      '' British
                    FdtLdt = ("29/" & Format(MonthByt, "00") & "/" & CurYear)
                End Select
            Else
                Select Case bytDateF
                Case 1      '' American
                    FdtLdt = (Format(MonthByt, "00") & "/28/" & CurYear)
                Case 2      '' British
                    FdtLdt = ("28/" & Format(MonthByt, "00") & "/" & CurYear)
            End Select
            End If
    End Select
End If
Exit Function
ERR_P:
    ShowError ("First and Last Date :: Reports")
End Function


Public Function DlyRepName(ByVal bytDly As Byte) As Boolean
On Error GoTo DRepErr
DlyRepName = True
    Select Case typOptIdx.bytDly
        Case 0: 'Physical Arrival
            Set Report = crxApp.OpenReport(App.path & "\Reports\dlyArrival.rpt", 1)
          Report.FormulaFields.GetItemByName("Header").Text = "'Daily Arrival Report for the Date of  " & typRep.strDlyDate & " '"
          frmCRV.Caption = "Daily Arrival Report for the Date of  " & typRep.strDlyDate
        Case 1: 'Absent
          Set Report = crxApp.OpenReport(App.path & "\Reports\DlyAbsent.rpt", 1)
          Report.FormulaFields.GetItemByName("Header").Text = "'Daily Absent Report for the Date of    " & typRep.strDlyDate & " '"
          frmCRV.Caption = "Daily Absent Report for the Date of  " & typRep.strDlyDate
        'Case 2: 'Cont Absent
            'Set repname = ContAbsent: strRepName = "DCntiAbsent":
            'Set RsName = DELOG.rsWeekReport_Grouping
        Case 3: 'Late Arrival
            Set Report = crxApp.OpenReport(App.path & "\Reports\DlyLArrival.rpt", 1)
            Report.FormulaFields.GetItemByName("Header").Text = "'Daily Late Arrival Report for the Date of    " & typRep.strDlyDate & " '"
            frmCRV.Caption = "Daily Late Arrival Report for the Date of  " & typRep.strDlyDate
        Case 4: 'Early Dep
            Set Report = crxApp.OpenReport(App.path & "\Reports\dlyearlyDep.rpt", 1)
            Report.FormulaFields.GetItemByName("Header").Text = "'Daily Early Departure Report for the Date of    " & typRep.strDlyDate & " '"
            frmCRV.Caption = "Daily Early Departure Report for the Date of  " & typRep.strDlyDate
        Case 5: 'Perf
             Set Report = crxApp.OpenReport(App.path & "\Reports\DlyPerf.rpt", 1)
             Report.FormulaFields.GetItemByName("Header").Text = "'Daily  Peformance Report for the Date of    " & typRep.strDlyDate & " '"
             frmCRV.Caption = "Daily Peformance Report for the Date of  " & typRep.strDlyDate
        Case 6: 'Irreg
             Set Report = crxApp.OpenReport(App.path & "\Reports\DlyIrre.rpt", 1)
             Report.FormulaFields.GetItemByName("Header").Text = "'Daily Irregular  Report for the Date of    " & typRep.strDlyDate & " '"
             frmCRV.Caption = "Daily Irregular  Peformance Report for the Date of  " & typRep.strDlyDate
        Case 7: 'authorized / Unauthorized OverTime
                Set Report = crxApp.OpenReport(App.path & "\Reports\dlyauthorizedot.rpt", 1)
                Report.FormulaFields.GetItemByName("Header").Text = "'Daily Authorized Overtime Report for the Date of    " & typRep.strDlyDate & " '"
        Case 13: 'un authorized OT
              Set Report = crxApp.OpenReport(App.path & "\Reports\dlyunauthorizedot.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Daily UnAuthorized Report for the Date of   " & typRep.strDlyDate & " '"
              frmCRV.Caption = "Daily UnAuthorized   Report for the Date of  " & typRep.strDlyDate
        Case 8: 'Entries
              Set Report = crxApp.OpenReport(App.path & "\Reports\dlyentries.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Daily Entries  Report for the Date of  " & typRep.strDlyDate & " '"
              frmCRV.Caption = "Daily Entries   Report for the Date of  " & typRep.strDlyDate
        Case 9: 'Shift Arrangement
              Set Report = crxApp.OpenReport(App.path & "\Reports\dlyshiftarrangement.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Daily Shift Arrangement Report for the Date of   " & typRep.strDlyDate & " '"
              frmCRV.Caption = "Daily Shift Arrangement  Report for the Date of  " & typRep.strDlyDate
        Case 10: 'Manpower
              Set Report = crxApp.OpenReport(App.path & "\Reports\dlymanpower.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Daily Manpower Report for the Date of   " & typRep.strDlyDate & " '"
              frmCRV.Caption = "Daily Manpower  Report for the Date of  " & typRep.strDlyDate
        Case 11: 'OutDoor
              Set Report = crxApp.OpenReport(App.path & "\Reports\dlyoutdoorduty.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Daily OutDoor Report for the Date of   " & typRep.strDlyDate & " '"
              frmCRV.Caption = "Daily OutDooor  Report for the Date of  " & typRep.strDlyDate
        Case 12: 'Summary
              Set Report = crxApp.OpenReport(App.path & "\Reports\dlySummary.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Daily Summary Report for the Date of   " & typRep.strDlyDate & " '"
              frmCRV.Caption = "Daily Summary  Report for the Date of  " & typRep.strDlyDate
       
        End Select
        Report.FormulaFields.GetItemByName("Cname").Text = "'" & strCName & "'"
Exit Function
DRepErr:   ShowError ("Set Daily Report Name :: Reports")
    DlyRepName = False
End Function
Public Function SetRepName() As Boolean
 On Error GoTo RepErr
SetRepName = True
bytPoLa = 1             '' SETS REPORT ORIENTATION 1:DEFAULT PRINTER ORIENTATION  2:LANDSCAPE
Select Case bytRepMode
 Case 1  '' Daily
    If DlyRepName(typOptIdx.bytDly) = False Then
        SetRepName = False
        Exit Function
    End If
Case 2  'Weekly
        Select Case typOptIdx.bytWek
        Case 0: '**Performance Report ***
                Set Report = crxApp.OpenReport(App.path & "\Reports\wkperformance.rpt", 1)
               Report.FormulaFields.GetItemByName("Header").Text = "'Weekly Performance Report for the week beginning from the date   " & typRep.strWkDate & " ' "
              frmCRV.Caption = "Weekly Performance Report for the Week beginning from the Date   " & typRep.strWkDate
        Case 1: '***ABSENT REPORT ****
             Set Report = crxApp.OpenReport(App.path & "\Reports\WkAbsent.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Weekly Absent Report for the week beginning from the date " & typRep.strWkDate & " '"
              Report.FormulaFields.GetItemByName("D1").Text = "'" & Left(WeekdayName(WeekDay(typRep.strWkDate, vbUseSystemDayOfWeek)), 3) & " '"
              Report.FormulaFields.GetItemByName("D2").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 1, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D3").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 2, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D4").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 3, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D5").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 4, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D6").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 5, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D7").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 6, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("Dt1").Text = "'" & Day(CDate(typRep.strWkDate)) & "'"
              Report.FormulaFields.GetItemByName("Dt2").Text = "'" & Day(CDate(typRep.strWkDate) + 1) & "'"
              Report.FormulaFields.GetItemByName("Dt3").Text = "'" & Day(CDate(typRep.strWkDate) + 2) & "'"
              Report.FormulaFields.GetItemByName("Dt4").Text = "'" & Day(CDate(typRep.strWkDate) + 3) & "'"
              Report.FormulaFields.GetItemByName("Dt5").Text = "'" & Day(CDate(typRep.strWkDate) + 4) & "'"
              Report.FormulaFields.GetItemByName("Dt6").Text = "'" & Day(CDate(typRep.strWkDate) + 5) & "'"
              Report.FormulaFields.GetItemByName("Dt7").Text = "'" & Day(CDate(typRep.strWkDate) + 6) & "'"
              frmCRV.Caption = "Weekly Absent Reportfor the Week beginning from the Date " & typRep.strWkDate
        Case 2: '***Attendance REPORT ***
              Set Report = crxApp.OpenReport(App.path & "\Reports\wkattendance.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "' Weekly Attendance Report for the week beginning from the date " & typRep.strWkDate & " ' "
              Report.FormulaFields.GetItemByName("D1").Text = "'" & Left(WeekdayName(WeekDay(typRep.strWkDate, vbUseSystemDayOfWeek)), 3) & " '"
              Report.FormulaFields.GetItemByName("D2").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 1, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D3").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 2, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D4").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 3, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D5").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 4, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D6").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 5, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D7").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 6, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("Dt1").Text = "'" & Day(CDate(typRep.strWkDate)) & "'"
              Report.FormulaFields.GetItemByName("Dt2").Text = "'" & Day(CDate(typRep.strWkDate) + 1) & "'"
              Report.FormulaFields.GetItemByName("Dt3").Text = "'" & Day(CDate(typRep.strWkDate) + 2) & "'"
              Report.FormulaFields.GetItemByName("Dt4").Text = "'" & Day(CDate(typRep.strWkDate) + 3) & "'"
              Report.FormulaFields.GetItemByName("Dt5").Text = "'" & Day(CDate(typRep.strWkDate) + 4) & "'"
              Report.FormulaFields.GetItemByName("Dt6").Text = "'" & Day(CDate(typRep.strWkDate) + 5) & "'"
              Report.FormulaFields.GetItemByName("Dt7").Text = "'" & Day(CDate(typRep.strWkDate) + 6) & "'"
              frmCRV.Caption = "Weekly Attendance Report for the Week beginning from the Date   " & typRep.strWkDate
        Case 3: 'Late Arrival
              Set Report = crxApp.OpenReport(App.path & "\Reports\wklatearrival.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Weekly Late Arrival Report for the Week beginning from the Date   " & typRep.strWkDate & " '"
              Report.FormulaFields.GetItemByName("D1").Text = "'" & Left(WeekdayName(WeekDay(typRep.strWkDate, vbUseSystemDayOfWeek)), 3) & " '"
              Report.FormulaFields.GetItemByName("D2").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 1, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D3").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 2, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D4").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 3, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D5").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 4, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D6").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 5, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D7").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 6, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("Dt1").Text = "'" & Day(CDate(typRep.strWkDate)) & "'"
              Report.FormulaFields.GetItemByName("Dt2").Text = "'" & Day(CDate(typRep.strWkDate) + 1) & "'"
              Report.FormulaFields.GetItemByName("Dt3").Text = "'" & Day(CDate(typRep.strWkDate) + 2) & "'"
              Report.FormulaFields.GetItemByName("Dt4").Text = "'" & Day(CDate(typRep.strWkDate) + 3) & "'"
              Report.FormulaFields.GetItemByName("Dt5").Text = "'" & Day(CDate(typRep.strWkDate) + 4) & "'"
              Report.FormulaFields.GetItemByName("Dt6").Text = "'" & Day(CDate(typRep.strWkDate) + 5) & "'"
              Report.FormulaFields.GetItemByName("Dt7").Text = "'" & Day(CDate(typRep.strWkDate) + 6) & "'"
              frmCRV.Caption = "Weekly Late Arrival Report for the Week beginning from the Date  " & typRep.strWkDate
        Case 4: 'Early Departure
             Set Report = crxApp.OpenReport(App.path & "\Reports\wkearlydeparture.rpt", 1)
             Report.FormulaFields.GetItemByName("Header").Text = "'Weekly Early Departure Report for the Week beginning from the Date   " & typRep.strWkDate & " '"
             Report.FormulaFields.GetItemByName("D1").Text = "'" & Left(WeekdayName(WeekDay(typRep.strWkDate, vbUseSystemDayOfWeek)), 3) & " '"
             Report.FormulaFields.GetItemByName("D2").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 1, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
             Report.FormulaFields.GetItemByName("D3").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 2, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
             Report.FormulaFields.GetItemByName("D4").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 3, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
             Report.FormulaFields.GetItemByName("D5").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 4, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
             Report.FormulaFields.GetItemByName("D6").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 5, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
             Report.FormulaFields.GetItemByName("D7").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 6, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
             Report.FormulaFields.GetItemByName("Dt1").Text = "'" & Day(CDate(typRep.strWkDate)) & "'"
             Report.FormulaFields.GetItemByName("Dt2").Text = "'" & Day(CDate(typRep.strWkDate) + 1) & "'"
             Report.FormulaFields.GetItemByName("Dt3").Text = "'" & Day(CDate(typRep.strWkDate) + 2) & "'"
             Report.FormulaFields.GetItemByName("Dt4").Text = "'" & Day(CDate(typRep.strWkDate) + 3) & "'"
             Report.FormulaFields.GetItemByName("Dt5").Text = "'" & Day(CDate(typRep.strWkDate) + 4) & "'"
             Report.FormulaFields.GetItemByName("Dt6").Text = "'" & Day(CDate(typRep.strWkDate) + 5) & "'"
             Report.FormulaFields.GetItemByName("Dt7").Text = "'" & Day(CDate(typRep.strWkDate) + 6) & "'"
             frmCRV.Caption = "Weekly Early Departure Report for the Week beginning from the Date   " & typRep.strWkDate
        Case 5: 'overtime
              Set Report = crxApp.OpenReport(App.path & "\Reports\wkovertime.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Weekly Overtime Report for the Week beginning from the Date   " & typRep.strWkDate & " '"
              Report.FormulaFields.GetItemByName("D1").Text = "'" & Left(WeekdayName(WeekDay(typRep.strWkDate, vbUseSystemDayOfWeek)), 3) & " '"
              Report.FormulaFields.GetItemByName("D2").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 1, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D3").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 2, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D4").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 3, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D5").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 4, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D6").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 5, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D7").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 6, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("Dt1").Text = "'" & Day(CDate(typRep.strWkDate)) & "'"
              Report.FormulaFields.GetItemByName("Dt2").Text = "'" & Day(CDate(typRep.strWkDate) + 1) & "'"
              Report.FormulaFields.GetItemByName("Dt3").Text = "'" & Day(CDate(typRep.strWkDate) + 2) & "'"
              Report.FormulaFields.GetItemByName("Dt4").Text = "'" & Day(CDate(typRep.strWkDate) + 3) & "'"
              Report.FormulaFields.GetItemByName("Dt5").Text = "'" & Day(CDate(typRep.strWkDate) + 4) & "'"
              Report.FormulaFields.GetItemByName("Dt6").Text = "'" & Day(CDate(typRep.strWkDate) + 5) & "'"
              Report.FormulaFields.GetItemByName("Dt7").Text = "'" & Day(CDate(typRep.strWkDate) + 6) & "'"
              frmCRV.Caption = "Weekly Overtime Report for the Week beginning from the Date   " & typRep.strWkDate
        Case 6: 'Shift Schedule
              Set Report = crxApp.OpenReport(App.path & "\Reports\wkshiftschedule.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Weekly Shift Schedule Report for the Week beginning from the Date    " & typRep.strWkDate & " '"
              Report.FormulaFields.GetItemByName("D1").Text = "'" & Left(WeekdayName(WeekDay(typRep.strWkDate, vbUseSystemDayOfWeek)), 3) & " '"
              Report.FormulaFields.GetItemByName("D2").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 1, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D3").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 2, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D4").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 3, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D5").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 4, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D6").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 5, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("D7").Text = "'" & Left(WeekdayName(WeekDay(DateAdd("d", 6, typRep.strWkDate), vbUseSystemDayOfWeek)), 3) & "'"
              Report.FormulaFields.GetItemByName("Dt1").Text = "'" & Day(CDate(typRep.strWkDate)) & "'"
              Report.FormulaFields.GetItemByName("Dt2").Text = "'" & Day(CDate(typRep.strWkDate) + 1) & "'"
              Report.FormulaFields.GetItemByName("Dt3").Text = "'" & Day(CDate(typRep.strWkDate) + 2) & "'"
              Report.FormulaFields.GetItemByName("Dt4").Text = "'" & Day(CDate(typRep.strWkDate) + 3) & "'"
              Report.FormulaFields.GetItemByName("Dt5").Text = "'" & Day(CDate(typRep.strWkDate) + 4) & "'"
              Report.FormulaFields.GetItemByName("Dt6").Text = "'" & Day(CDate(typRep.strWkDate) + 5) & "'"
              Report.FormulaFields.GetItemByName("Dt7").Text = "'" & Day(CDate(typRep.strWkDate) + 6) & "'"
              frmCRV.Caption = "Weekly Shift Schedule Report for the Week beginning from the Date   " & typRep.strWkDate
        Case 7: 'Irregular
              Set Report = crxApp.OpenReport(App.path & "\Reports\wkirregular.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Weekly Irregular Report for the Week beginning from the Date   " & typRep.strWkDate & " '"
              frmCRV.Caption = "Weekly Irregular Report for the Week beginning from the Date   " & typRep.strWkDate
     
        End Select
        Report.FormulaFields.GetItemByName("Cname").Text = "'" & strCName & "'"
Case 3  'Monthly
    Select Case typOptIdx.bytMon
    Case 0

                Set Report = crxApp.OpenReport(App.path & "\Reports\monperf.rpt", 1)
                Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Performance Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
                frmCRV.Caption = "Monthly Performance Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear

             If strlstdt = "28" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "' '"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "29" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "30" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "''"
               ElseIf strlstdt = "31" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
            End If
        Case 1: 'Attendance

                Set Report = crxApp.OpenReport(App.path & "\Reports\monattend.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Attendance Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Attendance Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              Erase StrLvCD: Erase strAlv
                If strlstdt = "28" Then
                Report.FormulaFields.GetItemByName("dt77").Text = "' '"
                Report.FormulaFields.GetItemByName("dt78").Text = "' '"
                Report.FormulaFields.GetItemByName("dt79").Text = "' '"
                ElseIf strlstdt = "29" Then
                Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
                Report.FormulaFields.GetItemByName("dt78").Text = "' '"
                Report.FormulaFields.GetItemByName("dt79").Text = "' '"
                ElseIf strlstdt = "30" Then
                Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
                Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
                Report.FormulaFields.GetItemByName("dt79").Text = "' '"
                ElseIf strlstdt = "31" Then
                Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
                Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
                Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
                End If
        Case 2: 'Muster Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MonMuster.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Muster Report for the month of " & typRep.strMonMth & " " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Muster Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              If strlstdt = "28" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "' '"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "29" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "30" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "31" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
               End If
        Case 3: 'Monthly Presnt
              Set Report = crxApp.OpenReport(App.path & "\Reports\monmonthlypresent.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Present Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Present Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              If strlstdt = "28" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "' '"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "29" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "30" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "31" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
               End If
        Case 4: 'Monthly Absent
              Set Report = crxApp.OpenReport(App.path & "\Reports\monmonthlyabsent.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Absent Report for the Month of " & typRep.strMonMth & "   " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Absent Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              If strlstdt = "28" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "' '"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "29" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "30" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "31" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
               End If
        Case 5: 'Overtime
                    Set Report = crxApp.OpenReport(App.path & "\Reports\monperf.rpt", 1)
                    Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Overtime Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
                    frmCRV.Caption = "Monthly Overtime Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              '***
        Case 6
                Set Report = crxApp.OpenReport(App.path & "\Reports\monovertimepaid.rpt", 1)
                Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Overtime Paid Report for the Month of " & typRep.strMonMth & "   " & typRep.strMonYear & "' "
                frmCRV.Caption = "Monthly Overtime Paid Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              
        Case 7: 'Absent / Late  / Early Departure Memo
              Set Report = crxApp.OpenReport(App.path & "\Reports\monabsentmemo.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Absent Memo Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Absent Memo Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              Report.FormulaFields.GetItemByName("Fnmemo").Text = "'" & strCapSND & "'"
        Case 8: 'Absent/Late/Early
              Set Report = crxApp.OpenReport(App.path & "\Reports\monlateearlyabsent.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Absent\Late\Early Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Absent\Late\Early Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
        Case 9: 'Leave Balance
              Set Report = crxApp.OpenReport(App.path & "\Reports\monleavebalance.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Leave Balance Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Leave Balance Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              If typDlyLvBal.bytDtOpt = 1 Then            '
                Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Leave Balance Report upto dated " & typDlyLvBal.DailyDt & "' "
                frmCRV.Caption = "Monthly Leave Balance Report upto dated " & typDlyLvBal.DailyDt & "' "
              End If
              Erase StrLvCD: Erase strAlv
        Case 10: 'Late Arrival
              Set Report = crxApp.OpenReport(App.path & "\Reports\monlt.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Late Arrival Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Late Arrival Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              If strlstdt = "28" Then
              Report.FormulaFields.GetItemByName("dt77").Text = "' '"
              Report.FormulaFields.GetItemByName("dt78").Text = "' '"
              Report.FormulaFields.GetItemByName("dt79").Text = "' '"
              ElseIf strlstdt = "29" Then
              Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
              Report.FormulaFields.GetItemByName("dt78").Text = "' '"
              Report.FormulaFields.GetItemByName("dt79").Text = "' '"
              ElseIf strlstdt = "30" Then
                    Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
                    Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
                   Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "31" Then
                   Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
                   Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
                   Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
               End If
        Case 11: 'Early Departure
              Set Report = crxApp.OpenReport(App.path & "\Reports\monel.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Early Departure Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
               frmCRV.Caption = "Monthly Early Departure Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
               If strlstdt = "28" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "' '"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "29" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "30" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "31" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
               End If
       Case 12: 'Late Memo
              Set Report = crxApp.OpenReport(App.path & "\Reports\monlatearrivalmemo.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Late Memo Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Late Memo Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
                 Report.FormulaFields.GetItemByName("Fnmemo").Text = "'" & strCapSND & "'"
      Case 13: 'Early Memo
              Set Report = crxApp.OpenReport(App.path & "\Reports\monearlydeparturememo.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Early Memo Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Early Memo Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              Report.FormulaFields.GetItemByName("Fnmemo").Text = "'" & strCapSND & "'"
        Case 14: 'Leave Consumption
              Set Report = crxApp.OpenReport(App.path & "\Reports\monleavconmp.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Leave Consumption Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Leave Consumption Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
        Case 15: 'Total Lates
              Set Report = crxApp.OpenReport(App.path & "\Reports\montotallates.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Total Lates Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Total Lates Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
              If GetFlagStatus("CombineLateEarly") Then
                frmCRV.Caption = "Monthly Total Lates-Early Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
                Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Total Lates-Early Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              End If
        Case 16: 'Total Earlys
              Set Report = crxApp.OpenReport(App.path & "\Reports\montotalearlys.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Total Earlys Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Total Earlys Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
       Case 17: 'Shift Schedule
              Set Report = crxApp.OpenReport(App.path & "\Reports\monshiftschedule.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Shift Schedule Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Shift Schedule Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
               If strlstdt = "28" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "' '"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "29" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "' '"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "30" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "' '"
               ElseIf strlstdt = "31" Then
               Report.FormulaFields.GetItemByName("dt77").Text = "'29'"
               Report.FormulaFields.GetItemByName("dt78").Text = "'30'"
               Report.FormulaFields.GetItemByName("dt79").Text = "'31'"
               End If
        Case 18: 'WO on Holiday
              Set Report = crxApp.OpenReport(App.path & "\Reports\monwohl.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Monthly Week off on Holiday Report for the Month of " & typRep.strMonMth & "  " & typRep.strMonYear & "' "
              frmCRV.Caption = "Monthly Week off on Holiday Report for the Month of " & typRep.strMonMth & " " & typRep.strMonYear
 
    End Select
           Report.FormulaFields.GetItemByName("Cname").Text = "'" & strCName & "'"
Case 4  'Yearly
    Select Case typOptIdx.bytYer
        Case 0: 'Absent
            ' Set Report = New YrAbPr
              Set Report = crxApp.OpenReport(App.path & "\Reports\yrabsent.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Yearly Absent Report for the Year of " & typRep.strYear & "'"
              frmCRV.Caption = "Yearly Absent Report for the Year of " & typRep.strYear
        Case 1: 'Mandays
             Set Report = crxApp.OpenReport(App.path & "\Reports\yrmandays.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Yearly Mandays Report for the Year of " & typRep.strYear & "'"
              Report.FormulaFields.GetItemByName("LVSTR").Text = "'" & YearStr & "'"
              frmCRV.Caption = "Yearly Mandays Report for the Year of " & typRep.strYear
              Erase StrLvCD: Erase strAlv
        Case 2: 'Performance
              Set Report = crxApp.OpenReport(App.path & "\Reports\yrperformance.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Yearly Performance Report for the Year of " & typRep.strYear & "'"
              
              Report.FormulaFields.GetItemByName("LVSTR").Text = "'" & YearStr & "'"
              frmCRV.Caption = "Yearly Performance Report for the Year of " & typRep.strYear
              Erase StrLvCD: Erase strAlv
        Case 3: 'Present
              Set Report = crxApp.OpenReport(App.path & "\Reports\yrpresent.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Yearly Present Report for the Year of " & typRep.strYear & "'"
              If typMaleFemale.femaleopt = 1 Then
              frmCRV.Caption = "Yearly Present Female for the Year of " & typRep.strYear
              ElseIf typMaleFemale.maleopt = 1 Then
              frmCRV.Caption = "Yearly Present Male for the Year of " & typRep.strYear
              Else
              frmCRV.Caption = "Yearly Present Report for the Year of " & typRep.strYear
              End If
        Case 4: 'Leave Information
                Set Report = crxApp.OpenReport(App.path & "\Reports\yrleaveinformation.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Yearly Leave Information Report for the Year of " & typRep.strYear & "'"
              frmCRV.Caption = "Yearly Leave Information Report for the Year of " & typRep.strYear
        Case 5: 'Leave Balance
              Set Report = crxApp.OpenReport(App.path & "\Reports\yrLeaveBlance.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Yearly Leave Balance Report for the Year of " & typRep.strYear & "'"
              frmCRV.Caption = "Yearly Leave Balance Report for the Year of " & typRep.strYear
        
 End Select
       Report.FormulaFields.GetItemByName("Cname").Text = "'" & strCName & "'"
Case 5  'Masters
    Select Case typOptIdx.bytMst
        Case 0: 'Employee List Report

               Set Report = crxApp.OpenReport(App.path & "\Reports\MstEmpList.rpt", 1)
         
              Report.FormulaFields.GetItemByName("Header").Text = "'Employee List Report '"
              frmCRV.Caption = "Employee List Report"
        Case 1: 'Employee Details Report
    
              Set Report = crxApp.OpenReport(App.path & "\Reports\mstempdetails.rpt", 1)
          
              Report.FormulaFields.GetItemByName("Header").Text = "'Employee Details Report '"
              frmCRV.Caption = "Employee Details Report"
        Case 2: 'Left Employee Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstLeftEmp.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Left Employee(s) Report for the Date of " & typRep.strLeftFr & " And  " & typRep.strLeftTo & "'"
              frmCRV.Caption = "Employee Details Report"
        Case 3: 'Leave Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstLeave.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Leave Master '"
              frmCRV.Caption = "Leave Details Report"
        Case 4: 'Shift Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstShift.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Shift Master Report '"
              frmCRV.Caption = "Shift Master Report"
        Case 5: 'Rotational Shift Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\mstrotshft.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Rotation Shift Master Report '"
              frmCRV.Caption = "Rotation Shift Master Report"
        Case 6: 'Holiday Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstHoliday.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Holiday Master Report '"
              frmCRV.Caption = "Holiday Master Report"
        Case 7: 'Department Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstDept.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Department Master Report'"
              frmCRV.Caption = "Department Master Report"
        Case 8: 'Category Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstCat.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Category Master Report'"
              frmCRV.Caption = "Category Master Report"
        Case 9: 'Group Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstGroup.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Group Master Report '"
              frmCRV.Caption = "Group Master Report"
        Case 10 'Location Report
              Set Report = crxApp.OpenReport(App.path & "\Reports\mstDesignation.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Designation Master Report '"
              frmCRV.Caption = "Location Master Report"
        Case 11 'Company
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstCompany.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Company Master Report '"
              frmCRV.Caption = "Company Master Report"
        Case 12 'Division
              Set Report = crxApp.OpenReport(App.path & "\Reports\MstDivision.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Division Master Report'"
              frmCRV.Caption = "Division Master Report"
        End Select
        Report.FormulaFields.GetItemByName("Cname").Text = "'" & strCName & "'"
Case 6  'Periodic
    Select Case typOptIdx.bytPer
        Case 0: 'Performance
              Set Report = crxApp.OpenReport(App.path & "\Reports\prperformance.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Performance Report for the period of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
              frmCRV.Caption = "Periodic Performance Report for the period from " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
        Case 1: 'Muster
              Set Report = crxApp.OpenReport(App.path & "\Reports\prmuster.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Muster Report for the period of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
              frmCRV.Caption = "Periodic Muster Report for the Period From " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
        Case 2
                Set Report = crxApp.OpenReport(App.path & "\Reports\provertime.rpt", 1)
                Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Overtime Report for the period of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
                frmCRV.Caption = "Periodic Overtime Report for the Period From " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
            
        Case 3: 'Late Arrival
              Set Report = crxApp.OpenReport(App.path & "\Reports\peLT.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Late Arrival Report for the Period of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
              frmCRV.Caption = "Periodic Late Arrival Report for the Period From " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
        Case 4: 'Early Departure
              Set Report = crxApp.OpenReport(App.path & "\Reports\peEL.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Early Departure Report for the Period of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
              frmCRV.Caption = "Periodic Early Departure Report for the Period From " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
        Case 5

              Set Report = crxApp.OpenReport(App.path & "\Reports\prconabsent.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Continuous Absent Report for the Period of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
              frmCRV.Caption = "Periodic Continuous Present Report for the Period From " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
        Case 6: 'Summary

                Set Report = crxApp.OpenReport(App.path & "\Reports\peSummary.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Summary Report for the period Of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
              frmCRV.Caption = "Periodic Summary Report for the period Of " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
        Case 7: 'Attendance
              Set Report = crxApp.OpenReport(App.path & "\Reports\prAttendance.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Periodic Muster Report for the period of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
              frmCRV.Caption = "Periodic Muster Report for the Period From " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
 
        Case 16:
                Set Report = crxApp.OpenReport(App.path & "\Reports\prPhysicalAbsent.rpt", 1)
                Report.FormulaFields.GetItemByName("Header").Text = "'Periodic physical Absent Report for the period Of " & typRep.strPeriFr & " And " & typRep.strPeriTo & " '"
                frmCRV.Caption = "Periodic physical Absent Report for the period Of " & DateDisp(typRep.strPeriFr) & " To " & DateDisp(typRep.strPeriTo)
        Case 17: 'Performance
              Set Report = crxApp.OpenReport(App.path & "\Reports\prperformanceExp.rpt", 1)
              Report.FormulaFields.GetItemByName("Header").Text = "'Performance Data " & typRep.strPeriFr & " TO " & typRep.strPeriTo & "'"
              Report.FormulaFields.GetItemByName("Cname").Text = "''"
              frmCRV.Caption = "Periodic Performance Report for the period from " & DateDisp(typRep.strPeriFr) & " And " & DateDisp(typRep.strPeriTo)
        
   End Select
            Report.FormulaFields.GetItemByName("Cname").Text = "'" & strCName & "'"
  End Select
Call Rpt_Intialization
If blnIntz Then Set CRV = frmCRV.CRV: Call PrnAtul: CRV.ReportSource = Report: Call SetFormIcon(frmCRV)
Exit Function
RepErr:   ShowError ("Set Report Name :: Reports")
    SetRepName = False
'    Resume Next
End Function
Public Function WkCreateFiles() As Boolean
On Error GoTo ERR_P
WkCreateFiles = False
Select Case typOptIdx.bytWek
    Case 1, 2, 3, 4, 5, 6
        Call GetReportFile("WStat")
        Call CreRepFile("WStat")
        WkCreateFiles = True
    Case 0, 7
        Call GetReportFile("WPerf")
        Call CreRepFile("WPerf")
        WkCreateFiles = True
    Case Else
        MsgBox NewCaptionTxt("M7002", adrsMod), vbExclamation
        WkCreateFiles = False
End Select
Exit Function
ERR_P:
    ShowError ("Weekly Create Files :: Reports")
End Function
Public Function monCreateFiles() As Boolean
On Error GoTo ERR_P
monCreateFiles = False
Select Case typOptIdx.bytMon
Case 0, 5, 10, 11
         Call GetReportFile("MPerf")
         Call CreRepFile("MPerf")
         monCreateFiles = True
         monCreateFiles = True

    Case 1, 2, 3, 4, 7, 12, 13, 17
          Call GetReportFile("MonperfB", 1)
          Call CreRepFile("MonperfB", 1)
          monCreateFiles = True
    Case 9, 15, 16
        Call GetReportFile("MAtt")
        Call CreRepFile("MAtt")
        monCreateFiles = True
   Case 8, 14
        Call GetReportFile("MALE")
        Call CreRepFile("MALE")
        monCreateFiles = True
   Case 6, 18
        monCreateFiles = True
    
   Case Else
        MsgBox NewCaptionTxt("M7002", adrsMod), vbExclamation
        monCreateFiles = False
End Select
Exit Function
ERR_P:
    ShowError ("Monthly Create Files :: Reports")
End Function

Public Function yrCreateFiles() As Boolean
On Error GoTo ERR_P
yrCreateFiles = False
Select Case typOptIdx.bytYer
    Case 0, 3
        Call GetReportFile("YrAbPr")
        Call CreRepFile("YrAbPr")
        yrCreateFiles = True
    Case 1, 2, 4
         Call GetReportFile("YrTB")
         Call CreRepFile("YrTB")
         yrCreateFiles = True
    Case 5
        yrCreateFiles = True
    Case Else
        MsgBox NewCaptionTxt("M7002", adrsMod), vbExclamation
        yrCreateFiles = False
End Select
Exit Function
ERR_P:
    ShowError ("Yearly Create Files :: Reports")
End Function

Public Function peCreateFiles() As Boolean
On Error GoTo ERR_P
peCreateFiles = False
Select Case typOptIdx.bytPer
    Case 0, 2, 3, 4, 16, 17
         Call GetReportFile("Mperf")
         Call CreRepFile("Mperf")
         peCreateFiles = True
   Case 1, 7 ' muster Report
         Call GetReportFile("MonperfA")
         Call CreRepFile("MonperfA")
         Call GetReportFile("MonperfB", 1)
         Call CreRepFile("MonperfB", 1)
          peCreateFiles = True
    Case 5
        Call GetReportFile("MonperfB", 1)
        Call CreRepFile("MonperfB", 1)
        peCreateFiles = True
    Case 6  ''Summary
        Call GetReportFile("DSumC")
        Call CreRepFile("DSumC")
        peCreateFiles = True
    Case Else
        MsgBox NewCaptionTxt("M7002", adrsMod), vbExclamation
        peCreateFiles = False
End Select
Exit Function
ERR_P:
    ShowError ("Periodically Create Files :: Reports")
End Function



'*** ***
Public Function WKPerfOvt() As Boolean
On Error GoTo ERR_P
WKPerfOvt = True
Dim lsum As Single, esum As Single, wsum As Single, osum As Single
Dim A_Str As String, D_Str As String, L_Str As String, E_str As String
Dim W_Str As String, O_Str As String, p_str As String, S_Str As String
Dim strGP As String, strOvt As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim DTESTR As String, STRECODE As String
Dim strfile1 As String, strFile2 As String

'********
Dim arrDTE(7) As String
Dim arrstr(7) As String
Dim depStr(7) As String
Dim lateStr(7) As Single
Dim erlyStr(7) As Single
Dim otStr(7) As Single
Dim wrkHrs(7) As Single
Dim prsStr(7) As String
Dim shftStr(7) As String
Dim i As String
i = 1

dtFirstDate = DateCompDate(typRep.strWkDate)
dtfromdate = DateCompDate(typRep.strWkDate)
dttodate = DateCompDate(typRep.strWkDate) + 6

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

strOvt = ""
If typOptIdx.bytWek = 6 Then strOvt = " and ovtim>0 and OTConf ='Y' "

DTESTR = ""
arrDTE(7) = ""
arrstr(7) = ""
depStr(7) = ""
lateStr(7) = 0
erlyStr(7) = 0
otStr(7) = 0
wrkHrs(7) = 0
prsStr(7) = ""
shftStr(7) = ""


Do While dtfromdate <= dttodate
    If dtfromdate = dttodate Then
        arrDTE(i) = Day(dtfromdate)
    ElseIf dtfromdate <> dttodate Then
        arrDTE(i) = Day(dtfromdate) & Spaces(Len(Trim(str(Day(dtfromdate)))))
    End If
    dtfromdate = DateAdd("d", 1, dtfromdate)
    i = i + 1
Loop
If strfile1 = strFile2 Then

'    If typOptIdx.bytWek = 0 Then 'Performance
    If typOptIdx.bytWek = 0 Or typOptIdx.bytWek = 8 Then  'Grish 03-04

    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs, " & _
    "ovtim,presabs," & strfile1 & ".shift,OTConf from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & DateCompStr(dtFirstDate) & _
    strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(dttodate) & strDTEnc & strOvt & " " & strSql
    Else
    
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,actrt_o,actrt_i,time5,time6, " & _
    "presabs," & strfile1 & ".shift,OTConf,ovtim from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode AND " & strfile1 & ".entry In (1,3,5,7) and " & strKDate & ">=" & strDTEnc & DateCompStr(dtFirstDate) & _
    strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(dttodate) & strDTEnc & strOvt & " " & strSql & " " 'And " & strFile1 & ".chq='*' "
    End If
    
Else
 If typOptIdx.bytWek = 0 Then  'performance
       strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs," & _
    "ovtim,presabs," & strfile1 & ".shift,OTConf from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & DateCompStr(dtFirstDate) & _
    strDTEnc & strOvt & strSql & " union select " & strFile2 & ".Empcode," & strKDate & "," & _
    "arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim,presabs," & strFile2 & ".shift,OTConf from " & _
    strFile2 & "," & rpTables & " where " & strFile2 & ".Empcode = empmst.Empcode " & _
    "and " & strKDate & "<=" & strDTEnc & DateCompStr(dttodate) & strDTEnc & strOvt & " " & strSql
    
    Else            'irregular
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,actrt_i,actrt_o,time5," & _
    "time6,presabs," & strfile1 & ".shift,OTConf,ovtim, latehrs,earlhrs,wrkhrs from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode AND " & strfile1 & ".entry In (1,3,5,7) and " & strKDate & ">=" & strDTEnc & DateCompStr(dtFirstDate) & _
    strDTEnc & strOvt & strSql & _
    " union select " & strFile2 & ".Empcode," & strKDate & ",arrtim,deptim,actrt_i,actrt_o,time5," & _
    "time6,presabs," & strFile2 & ".shift,OTConf,ovtim, latehrs,earlhrs,wrkhrs from " & _
    strFile2 & "," & rpTables & " where " & strFile2 & ".Empcode = empmst.Empcode AND " & strFile2 & ".entry In (1,3,5,7) " & _
    "and " & strKDate & "<=" & strDTEnc & DateCompStr(dttodate) & strDTEnc & strOvt & " " & strSql & " And " & strFile2 & ".chq='*' "
    End If
   
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
i = 1
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
            STRECODE = adrsTemp!Empcode
            dtfromdate = dtFirstDate
            Do While dtfromdate <= dttodate
                If adrsTemp.EOF Then Exit Do
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtfromdate Then
                        arrstr(i) = IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), "0.00")
                                                        
                        depStr(i) = IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), "0.00")
                                                                                        
                                                                                       
                        If typOptIdx.bytWek = 0 Or typOptIdx.bytWek = 8 Then  'Grish 03-04
                        
                           lateStr(i) = IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                           Spaces(Len(Format(adrsTemp!latehrs, "0.00"))) & Format(adrsTemp!latehrs, "0.00"), "0.00")
                                                        
                            erlyStr(i) = IIf(Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                            Spaces(Len(Format(adrsTemp!earlhrs, "0.00"))) & Format(adrsTemp!earlhrs, "0.00"), "0.00")
                                                        
                            wrkHrs(i) = IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                            Spaces(Len(Format(adrsTemp!wrkHrs, "0.00"))) & Format(adrsTemp!wrkHrs, "0.00"), "0.00")
                        Else
                        lateStr(i) = IIf(Not IsNull(adrsTemp!actrt_i) And adrsTemp!actrt_i > 0, _
                        Spaces(Len(Format(adrsTemp!actrt_i, "0.00"))) & Format(adrsTemp!actrt_i, "0.00"), "0.00")
                                                        
                        erlyStr(i) = IIf(Not IsNull(adrsTemp!Actrt_O) And adrsTemp!Actrt_O > 0, _
                        Spaces(Len(Format(adrsTemp!Actrt_O, "0.00"))) & Format(adrsTemp!Actrt_O, "0.00"), "0.00")
                                                        
                        wrkHrs(i) = IIf(Not IsNull(adrsTemp!time5) And adrsTemp!time5 > 0, _
                        Spaces(Len(Format(adrsTemp!time5, "0.00"))) & Format(adrsTemp!time5, "0.00"), "0.00")
                        
                        
                        End If
                        
                        
                        
                        

                        If adrsTemp("OTConf") = "Y" Then     ''if authorized OT then only Calculate and show
                            otStr(i) = IIf(Not IsNull(adrsTemp!ovtim) And adrsTemp!ovtim > 0, _
                            Spaces(Len(Format(adrsTemp!ovtim, "0.00"))) & Format(adrsTemp!ovtim, "0.00"), "0.00")
                            osum = TimAdd(IIf(IsNull(osum), 0, osum), IIf(IsNull(adrsTemp!ovtim), 0, adrsTemp!ovtim))
                     Else
                       otStr(i) = 0
         
                        End If
                        
                        prsStr(i) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                                                        
                        shftStr(i) = IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "") & Spaces(Len(adrsTemp!Shift))
                        If UCase(Trim(prsStr(i))) = "WOWO" And (Val(arrstr(i)) > 0 Or Val(depStr(i)) > 0) Then prsStr(i) = Trim(prsStr(i)) & "p"
                        If typOptIdx.bytWek = 0 Then
                        lsum = TimAdd(IIf(IsNull(lsum), 0, lsum), IIf(IsNull(adrsTemp!latehrs) Or adrsTemp!latehrs <= 0, 0, adrsTemp!latehrs))
                        esum = TimAdd(IIf(IsNull(esum), 0, esum), IIf(IsNull(adrsTemp!earlhrs) Or adrsTemp!earlhrs <= 0, 0, adrsTemp!earlhrs))
                        wsum = TimAdd(IIf(IsNull(wsum), 0, wsum), IIf(IsNull(adrsTemp!wrkHrs), 0, adrsTemp!wrkHrs))
                        End If
                    ElseIf adrsTemp!Date <> dtfromdate Then
                        arrstr(i) = ""  ' 16-Nov
                        depStr(i) = ""
                        A_Str = A_Str & Spaces(0)
                        D_Str = D_Str & Spaces(0)
                        L_Str = L_Str & Spaces(0)
                        E_str = E_str & Spaces(0)
                        W_Str = W_Str & Spaces(0)
                        O_Str = O_Str & Spaces(0)
                        p_str = p_str & Spaces(0)
                        S_Str = S_Str & Spaces(0)
                    End If
                Else
                    Exit Do
                End If
                If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
                dtfromdate = DateAdd("d", 1, dtfromdate)
            i = i + 1
            If i > 7 Then Exit Do
            Loop 'END OF DATE LOOP
            ConMain.Execute "insert into " & strRepFile & "" & _
            "( Empcode,d1,d2,d3,d4,d5,d6,d7,arr1,arr2,arr3,arr4,arr5,arr6,arr7,dep1,dep2,dep3,dep4,dep5,dep6,dep7,late1,late2," & _
            "late3,late4,late5,late6,late7,erly1,erly2,erly3,erly4,erly5,erly6,erly7,OT1,OT2,OT3,OT4,OT5,OT6,OT7,wrk1," & _
            " wrk2 , wrk3, wrk4, wrk5, wrk6, wrk7, pres1, pres2, pres3, pres4, pres5, pres6, pres7, shft1, shft2, shft3, shft4, shft5, shft6, shft7 " & _
            ",sumextra)  values('" & STRECODE & _
            "','" & arrDTE(1) & "','" & arrDTE(2) & "','" & arrDTE(3) & "','" & arrDTE(4) & "','" & arrDTE(5) & "','" & arrDTE(6) & "','" & arrDTE(7) & "', " & _
            "'" & arrstr(1) & "','" & arrstr(2) & "','" & arrstr(3) & "','" & arrstr(4) & "','" & arrstr(5) & "','" & arrstr(6) & "','" & arrstr(7) & "', " & _
            " '" & depStr(1) & "','" & depStr(2) & "','" & depStr(3) & "','" & depStr(4) & "','" & depStr(5) & "','" & depStr(6) & "','" & depStr(7) & "'," & _
            lateStr(1) & "," & lateStr(2) & "," & lateStr(3) & "," & lateStr(4) & "," & lateStr(5) & "," & lateStr(6) & "," & lateStr(7) & "," & _
            erlyStr(1) & "," & erlyStr(2) & "," & erlyStr(3) & "," & erlyStr(4) & "," & erlyStr(5) & "," & erlyStr(6) & "," & erlyStr(7) & "," & _
            otStr(1) & "," & otStr(2) & "," & otStr(3) & "," & otStr(4) & "," & otStr(5) & "," & otStr(6) & "," & otStr(7) & "," & _
            wrkHrs(1) & "," & wrkHrs(2) & "," & wrkHrs(3) & "," & wrkHrs(4) & "," & wrkHrs(5) & "," & wrkHrs(6) & "," & wrkHrs(7) & "," & _
            "'" & prsStr(1) & "','" & prsStr(2) & "','" & prsStr(3) & "','" & prsStr(4) & "','" & prsStr(5) & "','" & prsStr(6) & "','" & prsStr(7) & "'," & _
            "'" & shftStr(1) & "','" & shftStr(2) & "','" & shftStr(3) & "','" & shftStr(4) & "','" & shftStr(5) & "','" & shftStr(6) & "','" & shftStr(7) & "'," & _
              osum & ")"

            
       '" '" & wrkHrs(1) & "','" & wrkHrs(2) & "','" & wrkHrs(3) & "','" & wrkHrs(4) & "','" & wrkHrs(5) & "','" & wrkHrs(6) & "','" & wrkHrs(7) & "', "
            lsum = 0: esum = 0: wsum = 0: osum = 0
            i = 1
            dtfromdate = dtFirstDate
        Loop 'END OF EMPLOYEE LOOP
    End If
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    WKPerfOvt = False
End If 'adrsTemp.eof
Exit Function
ERR_P:
    ShowError ("Periodic Performance Overtime :: Reports")
    WKPerfOvt = False
''    Resume Next
End Function

Public Function WkPerfo(strFrDate As String, strLtDate As String) As Boolean
On Error GoTo RepErr
WkPerfo = True
Dim DTESTR As String, dtTempDate As Date

dtTempDate = DateCompDate(strFrDate)
DTESTR = ""
Do While dtTempDate <= DateCompDate(strLtDate)
    If dtTempDate = DateCompDate(strLtDate) Then
        DTESTR = DTESTR & Day(dtTempDate)
    ElseIf dtTempDate <> strLtDate Then
        DTESTR = DTESTR & Day(dtTempDate) & Spaces(Len(Trim(str(Day(dtTempDate)))))
    End If
    dtTempDate = DateAdd("d", 1, dtTempDate)
Loop
dtTempDate = DateCompDate(strFrDate)

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open " select distinct Empcode from " & rpTables & " where Empcode = Empcode " & _
strSql & " order by Empcode", ConMain
If Not adrsEmp.BOF And Not adrsEmp.EOF Then
    Do While (Not adrsEmp.EOF)
        Call wkPerfCalc(adrsEmp!Empcode, dtTempDate, DTESTR)
        adrsEmp.MoveNext
    Loop
Else
    Call SetMSF1Cap(10)
    WkPerfo = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If
If adrsEmp.State = 1 Then adrsEmp.Close
Exit Function
RepErr:
    ShowError ("Weekly Performance :: Reports")
    WkPerfo = False
End Function

Private Sub wkPerfCalc(ByVal Code As String, ByVal dateFrom As Date, ByVal strDate As String)
On Error GoTo ERR_P
Dim A_Str As String, D_Str As String, p_str As String, S_Str As String
Dim osum As Single, wsum As Single, lsum As Single, esum As Single
Dim L_Str As String, E_str As String, W_Str As String, O_Str As String
Dim PunStr As String, Back_str As String, dateTo As Date

lsum = 0: esum = 0: wsum = 0:   osum = 0
A_Str = "": D_Str = "": L_Str = "": E_str = ""
W_Str = "": O_Str = "": p_str = "": S_Str = ""
PunStr = "": Back_str = ""
 
dateTo = DateAdd("d", 6, dateFrom)

Do While dateFrom <= dateTo
    strMon_Trn = ""
    strMon_Trn = MakeName(MonthName(Month(dateFrom)), Year(dateFrom), "trn")
    If adrsTemp.State = 1 Then adrsTemp.Close
    If typOptIdx.bytWek = 0 Then  'performance
        adrsTemp.Open "select arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim,presabs,shift,OTConf from " & _
        strMon_Trn & " where " & strKDate & " =  " & strDTEnc & DateCompStr(CStr(dateFrom)) & _
        strDTEnc & " and Empcode = '" & Code & "'", ConMain, adOpenStatic
    Else            'irregular
        adrsTemp.Open "select arrtim,deptim,actrt_i,actrt_o,time5,time6,presabs,shift,OTConf " & _
        ",entry from " & strMon_Trn & " where " & strKDate & " =  " & strDTEnc & _
        DateCompStr(CStr(dateFrom)) & strDTEnc & " and Empcode = '" & Code & _
        "' and " & strMon_Trn & ".chq='*' ", ConMain, adOpenStatic
    End If
    If Not adrsTemp.EOF And Not adrsTemp.BOF Then
        A_Str = A_Str & wkstrDoit(0)            '"arrtim"
        D_Str = D_Str & wkstrDoit(1)            '"deptim"
        L_Str = L_Str & wkstrDoitRef(2, lsum)   '"latehrs"/actrt_i
        E_str = E_str & wkstrDoitRef(3, esum)   '"earlhrs"/actrt_o
        W_Str = W_Str & wkstrDoitRef(4, wsum)   '"wrkhrs"/time5
        O_Str = O_Str & wkstrDoitRef(5, osum)   '"ovtim"/time6
        p_str = p_str & wkstrDoit(6)            '"presabs"
        S_Str = S_Str & wkstrDoit(7)            '"shift"
        
        If typOptIdx.bytWek = 7 Then
            PunStr = PunStr & wkstrDoit(9)      '"entry"
        End If
    Else
        A_Str = A_Str & Spaces(0)
        D_Str = D_Str & Spaces(0)
        L_Str = L_Str & Spaces(0)
        E_str = E_str & Spaces(0)
        W_Str = W_Str & Spaces(0)
        O_Str = O_Str & Spaces(0)
        p_str = p_str & Spaces(0)
        S_Str = S_Str & Spaces(0)
        If typOptIdx.bytWek = 7 Then
            PunStr = PunStr & Spaces(0)      '"entry"
        End If
    End If
    dateFrom = DateAdd("d", 1, dateFrom)
Loop
If Trim(A_Str) = "" And Trim(D_Str) = "" And Trim(L_Str) = "" And Trim(E_str) _
= "" And Trim(W_Str) = "" And Trim(O_Str) = "" And Trim(p_str) = "" And _
Trim(S_Str) = "" Then
Else
    If lsum < 0 Then lsum = 0
    If esum < 0 Then esum = 0
    If wsum < 0 Then wsum = 0
    If osum < 0 Then osum = 0
    ConMain.Execute "insert into " & strRepFile & "(Empcode," & strKDate & ",arrstr," & _
    "depstr,latestr,earlstr,workstr,otstr,presabsstr,shfstr,sumlate,sumearly,sumwork," & _
    "sumextra,punches)  values('" & Code & "','" & strDate & "','" & A_Str & "','" & D_Str & "','" & _
    L_Str & "','" & E_str & "','" & W_Str & "','" & O_Str & "','" & p_str & "','" & S_Str & "'," & lsum & _
    "," & esum & "," & wsum & "," & osum & ",'" & PunStr & "')"
End If
Exit Sub
ERR_P:
    ShowError ("Weekly Performance Calc:: Reports")
End Sub

Public Function wkstrDoit(ByVal intIndex As Integer) As String
On Error GoTo ERR_P
If intIndex = 7 Then ''For shift Report
    wkstrDoit = IIf(adrsTemp(intIndex) > 0, _
      Spaces(Len(adrsTemp(intIndex))) & adrsTemp(intIndex), Spaces(0))
Else
    wkstrDoit = IIf(adrsTemp(intIndex) > 0, _
      Spaces(Len(Format(adrsTemp(intIndex), "0.00"))) & Format(adrsTemp(intIndex), "0.00"), Spaces(0))
End If
Exit Function
ERR_P:
    ShowError ("wkstrDoit :: Reports")
    wkstrDoit = Spaces(0)
End Function

Public Function wkstrDoitRef(ByVal intIndex As Integer, ByRef sngSum As Single) _
    As String
On Error GoTo ERR_P
If intIndex = 5 Then        ''if OT Field then
    If adrsTemp("OTConf") = "Y" Then ''if authorized then only calculate and show
        wkstrDoitRef = IIf(adrsTemp(intIndex) > 0, _
            Spaces(Len(Format(adrsTemp(intIndex), "0.00"))) & _
            Format(adrsTemp(intIndex), "0.00"), Spaces(0))
        sngSum = TimAdd(IIf(IsNull(sngSum), 0, Format(sngSum, "0.00")), _
            IIf(IsNull(adrsTemp(intIndex)), 0, Format(adrsTemp(intIndex), "0.00")))
    Else
        wkstrDoitRef = Spaces(0)
    End If
Else
    wkstrDoitRef = IIf(adrsTemp(intIndex) > 0, _
        Spaces(Len(Format(adrsTemp(intIndex), "0.00"))) & Format(adrsTemp(intIndex), _
        "0.00"), Spaces(0))
    sngSum = TimAdd(IIf(IsNull(sngSum), 0, Format(sngSum, "0.00")), IIf(IsNull(adrsTemp(intIndex)), 0, Format(adrsTemp(intIndex), "0.00")))
End If
Exit Function
ERR_P:
    ShowError ("wkstrDoit :: Reports")
    wkstrDoitRef = Spaces(0)
End Function

Public Function WkOtherRep() As Boolean
On Error GoTo RepErr
WkOtherRep = True
Dim dtfromdate As Date, dttodate As Date, i As Byte
Dim stropt As String, strfile1 As String, strFile2 As String
Dim strMonth As String, strYear As String
Dim strArrWeek(6) As String
strArrWeek(0) = "frw": strArrWeek(1) = "secw": strArrWeek(2) = "thw"
strArrWeek(3) = "fow": strArrWeek(4) = "fiw": strArrWeek(5) = "siw"
strArrWeek(6) = "sevw"
Select Case typOptIdx.bytWek
    Case 1: stropt = "presabs"
    Case 2: stropt = "presabs"
    Case 3: stropt = "latehrs"
    Case 4: stropt = "earlhrs"
    Case 5: stropt = "ovtim"
End Select
dtfromdate = DateCompDate(typRep.strWkDate)
dttodate = DateCompDate(typRep.strWkDate) + 6

If DateDiff("m", dtfromdate, dttodate) = 0 Then
    strMonth = MonthName(Month(dttodate))
    strYear = Year(DateCompDate(dttodate))
    strfile1 = MakeName(strMonth, strYear, "trn")
    strFile2 = strfile1
Else
    strMonth = MonthName(Month(DateCompDate(dtfromdate)))
    strYear = Year(DateCompDate(dtfromdate))
    strfile1 = MakeName(strMonth, strYear, "trn")          'first file
    strFile2 = MakeName(MonthName(Month(DateCompDate(dttodate))), _
    Year(DateCompDate(dttodate)), "trn")                   'second file
End If

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select distinct Empcode from " & rpTables & " where Empcode = Empcode " & _
strSql & " order by Empcode ", ConMain, adOpenStatic
Dim strGW As String
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    Do While Not adrsEmp.EOF
        If stropt = "presabs" Then
            strGW = "select Empcode," & stropt & "," & strKDate & ", arrtim, deptim  from " & strfile1 & " where " & strKDate & ">=" & _
            strDTEnc & DateCompStr(dtfromdate) & strDTEnc & " AND " & strKDate & "<=" & strDTEnc & _
            DateCompStr(dttodate) & strDTEnc & " And EmpCode = " & " '" & adrsEmp(0) & _
            "' union select Empcode," & stropt & "," & strKDate & ", arrtim, deptim  from " & strFile2 & " where " & strKDate & ">=" & _
            strDTEnc & DateCompStr(dtfromdate) & strDTEnc & " AND " & strKDate & "<=" & strDTEnc & _
            DateCompStr(dttodate) & strDTEnc & " and Empcode=" & "'" & _
            adrsEmp(0) & "'" & " order by Empcode," & strKDate & " "
        ElseIf stropt = "earlhrs" Or stropt = "latehrs" Or stropt = "ovtim" Then
            strGW = "select Empcode," & stropt & "," & strKDate & ",OTConf from " & strfile1 & " where " & strKDate & ">=" & _
            strDTEnc & DateCompStr(dtfromdate) & strDTEnc & " And " & strKDate & " <= " & strDTEnc & _
            DateCompStr(dttodate) & strDTEnc & "  And EmpCode = " & " '" & adrsEmp(0) & "' AND " & _
            stropt & ">0" & " union select Empcode," & stropt & "," & strKDate & ",OTConf from " & strFile2 & _
            " where " & strKDate & ">=" & strDTEnc & DateCompStr(dtfromdate) & strDTEnc & " AND " & strKDate & "<=" & _
            strDTEnc & DateCompStr(dttodate) & strDTEnc & " and Empcode=" & "'" & _
            adrsEmp(0) & "'" & " AND " & stropt & ">0" & "  order by Empcode," & strKDate & " "
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open strGW, ConMain, adOpenStatic
        If Not (adrsTemp.EOF And adrsTemp.BOF) Then
            With ConMain
                .Execute "delete  from " & strRepFile & " where Empcode='" & adrsEmp(0) & "'"
                .Execute "insert into " & strRepFile & " (Empcode) values (" & "'" & adrsEmp(0) & "'" & ")"
            End With
            For i = 1 To 7
                If adrsTemp.EOF = True Then Exit For
                If typOptIdx.bytWek = 1 Then
                    If UCase((Left(adrsTemp(1), 2))) <> pVStar.AbsCode And _
                    UCase((Right(adrsTemp(1), 2))) <> pVStar.AbsCode Then
                        GoTo skipInsert
                    End If
                End If
                Select Case Day(adrsTemp!Date)
                    Case Day(DateAdd("d", i - 1, dtfromdate))
                        If stropt = "ovtim" Then
                            ConMain.Execute "update " & strRepFile & " set " & _
                            strArrWeek(i - 1) & " =" & "'" & Format(adrsTemp(1), "0.00") & "' where " & _
                            "Empcode='" & adrsTemp!Empcode & "' and '" & adrsTemp("OTConf") & "' = 'Y'"
                        Else
                            sqlStr = "update " & strRepFile & " set " & _
                            strArrWeek(i - 1) & " =" & "'" & Format(adrsTemp(1), "0.00") & "' "
                            If stropt = "presabs" And UCase(Trim(adrsTemp(1))) = "WOWO" Then
                                If Val(adrsTemp("arrtim")) > 0 Or Val(adrsTemp("deptim")) > 0 Then sqlStr = Replace(sqlStr, "WOWO", "WOWOp")
                            End If
                            sqlStr = sqlStr + " where " & "Empcode='" & adrsTemp!Empcode & "' "
                            ConMain.Execute sqlStr
                        End If
                        adrsTemp.MoveNext
                End Select
                GoTo unskip
skipInsert:
                If Not adrsTemp.EOF Then adrsTemp.MoveNext
unskip:
            Next
        End If
        If Not adrsEmp.EOF Then adrsEmp.MoveNext
        'if bytbackend =2 then SLEEP (1000)
    Loop
Else
    Call SetMSF1Cap(10)
    WkOtherRep = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If
If adrsTemp.State = 1 Then adrsTemp.Close
ConMain.Execute "delete  from " & strRepFile & " where FrW is null and " & _
" SecW is null and thW is null and FoW is null and FiW is null and SiW is null and " & _
" SevW is null "
If adrsTemp.State = 1 Then adrsTemp.Close
If adrsEmp.State = 1 Then adrsEmp.Close
Exit Function
RepErr:
    ShowError ("Weekly Other Reports :: Reports")
    WkOtherRep = False
    'Resume Next
End Function

Public Function WkShiftRep() As Boolean
On Error GoTo RepErr
WkShiftRep = True
Dim bytDay As Byte, bytArrcnt As Byte
Dim dtfromdate As Date, dttodate As Date
Dim strShfFile1 As String, strShfFile2 As String
Dim strArrWeek(6) As String
strArrWeek(0) = "frw": strArrWeek(1) = "secw": strArrWeek(2) = "thw"
strArrWeek(3) = "fow": strArrWeek(4) = "fiw": strArrWeek(5) = "siw"
strArrWeek(6) = "sevw"

bytArrcnt = 0
dtfromdate = DateCompDate(typRep.strWkDate)
dttodate = DateCompDate(typRep.strWkDate) + 6

strShfFile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "shf")
If DateDiff("m", dtfromdate, dttodate) > 0 Then
    strShfFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "shf")
Else
    strShfFile2 = strShfFile1
End If

ConMain.Execute "insert into " & strRepFile & "(Empcode)" & _
" select distinct empmst.Empcode from " & rpTables & " ," & strShfFile1 & _
" where " & strShfFile1 & ".Empcode = empmst.Empcode " & strSql

Do While dtfromdate <= dttodate
    bytDay = Day(dtfromdate)
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select distinct " & strShfFile1 & ".Empcode," & "d" & bytDay & _
    " from " & strShfFile1 & "," & rpTables & " where " & strShfFile1 & _
    ".Empcode = empmst.Empcode " & strSql, ConMain, adOpenStatic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
        Do While Not adrsTemp.EOF
            ConMain.Execute " update " & strRepFile & " set " & strArrWeek(bytArrcnt) & _
            " = '" & IIf(IsEmpty(adrsTemp(1)) Or IsNull(adrsTemp(1)), Null, adrsTemp(1)) & "' where Empcode = '" & adrsTemp(0) & "'"
            adrsTemp.MoveNext
        Loop
    End If
    bytArrcnt = bytArrcnt + 1
    dtfromdate = DateAdd("d", 1, dtfromdate)
    If Month(dtfromdate) = Month(dttodate) Then strShfFile1 = strShfFile2
Loop

strShfFile1 = "frw is null and secw is null and thw is null and fow is null and " & _
"fiw is null and siw is null and sevw is null "
ConMain.Execute "delete from " & strRepFile & " where " & strShfFile1

strShfFile1 = "frw = '' and secw = '' and thw = '' and fow = '' and " & _
"fiw = '' and siw = '' and sevw = '' "
ConMain.Execute "delete from " & strRepFile & " where " & strShfFile1

If adrsTemp.State = 1 Then adrsTemp.Close
Exit Function
RepErr:
    ShowError ("Weekly Shift Report :: Reports")
    WkShiftRep = False
End Function

Public Function monPerfOt() As Boolean
On Error GoTo ERR_P
monPerfOt = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim osum As Single, lsum As Single, esum As Single, wsum As Single
Dim STRECODE As String, strTrnFile As String, strDateS As String
Dim A_Str As String, D_Str As String, L_Str As String, E_str As String
Dim W_Str As String, O_Str As String, p_str As String, S_Str As String
Dim strPatch As String
dtfromdate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "f"))
dttodate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))
strlstdt = Day(dttodate)
strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")
Select Case typOptIdx.bytMon
    Case 0: strPatch = ""
    Case 5: strPatch = " and " & strTrnFile & ".ovtim >0 and " & strTrnFile & ".OTConf = 'Y'"
    Case 10: strPatch = " and " & strTrnFile & ".latehrs >0 "
    Case 11: strPatch = " and " & strTrnFile & ".earLhrs >0 "
End Select

If adrsTemp.State = 1 Then adrsTemp.Close

        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & strPatch & " " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
    
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
'    strDateS = monDateStr(Day(DateCompStr((dttodate)))) 'Assigning date string
     strDateS = Format(dtfromdate, "dd/mmm/yyyy")
    Dim Flag As Integer
    Do While Not (adrsTemp.EOF) 'And dtTempDate < dtToDate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        A_Str = "": D_Str = "": L_Str = "": E_str = ""
        W_Str = "": O_Str = "": p_str = "": S_Str = ""
        lsum = 0: esum = 0: wsum = 0: osum = 0
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                        A_Str = A_Str & IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), Spaces(0))
                                       
                         D_Str = D_Str & IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), Spaces(0))
             
                    S_Str = S_Str & IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "") & Spaces(Len(adrsTemp!Shift))
                    
                    If typOptIdx.bytMon <> 11 Then
                        L_Str = L_Str & IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                            Spaces(Len(Format(adrsTemp!latehrs, "0.00"))) & Format(adrsTemp!latehrs, "0.00"), Spaces(0))
                        lsum = TimAdd(IIf(IsNull(lsum), 0, lsum), IIf(IsNull(adrsTemp!latehrs) Or adrsTemp!latehrs <= 0, 0, adrsTemp!latehrs))
                    End If
                    If typOptIdx.bytMon <> 10 Then
                        E_str = E_str & IIf(Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                             Spaces(Len(Format(adrsTemp!earlhrs, "0.00"))) & Format(adrsTemp!earlhrs, "0.00"), Spaces(0))
                        esum = TimAdd(IIf(IsNull(esum), 0, esum), IIf(IsNull(adrsTemp!earlhrs) Or adrsTemp!earlhrs <= 0, 0, adrsTemp!earlhrs))
                    End If
                    If typOptIdx.bytMon = 0 Or typOptIdx.bytMon = 5 Or typOptIdx.bytMon = 10 Or typOptIdx.bytMon = 11 Or typOptIdx.bytMon = 31 Or typOptIdx.bytMon = 47 Then '31 Add By Girish 27-04-2009
                        W_Str = W_Str & IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                        Spaces(Len(Format(adrsTemp!wrkHrs, "0.00"))) & Format(adrsTemp!wrkHrs, "0.00"), Spaces(0))
                        wsum = TimAdd(IIf(IsNull(wsum), 0, wsum), IIf(IsNull(adrsTemp!wrkHrs) Or adrsTemp!wrkHrs < 0, 0, adrsTemp!wrkHrs))
                        If typOptIdx.bytMon = 47 And Flag = 0 Then
                            O_Str = O_Str & IIf(Not IsNull(adrsTemp!LunchLtHrs) And adrsTemp!LunchLtHrs > 0, _
                             Spaces(Len(Format(adrsTemp!LunchLtHrs, "0.00"))) & Format(adrsTemp!LunchLtHrs, "0.00"), Spaces(0))
                            osum = TimAdd(IIf(IsNull(osum), 0, osum), IIf(IsNull(adrsTemp!LunchLtHrs), 0, adrsTemp!LunchLtHrs))
                        Else
                            If adrsTemp("OTConf") = "Y" Then
                                O_Str = O_Str & IIf(Not IsNull(adrsTemp!ovtim) And adrsTemp!ovtim > 0, _
                                Spaces(Len(Format(adrsTemp!ovtim, "0.00"))) & Format(adrsTemp!ovtim, "0.00"), Spaces(0))
                                osum = TimAdd(IIf(IsNull(osum), 0, osum), IIf(IsNull(adrsTemp!ovtim), 0, adrsTemp!ovtim))
                            Else
                                O_Str = O_Str & Spaces(0)
                            End If
                        End If
                                                    
                            If (Left(adrsTemp!presabs, 2) = pVStar.WosCode Or _
                            Left(adrsTemp!presabs, 2) = pVStar.HlsCode) And adrsTemp!arrtim > 0 Then
                                If Not (adrsTemp!presabs) = pVStar.HlsCode + pVStar.PrsCode Then 'Girish 2-07-09
                                    p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "     ") & _
                                    "p" & Spaces(Len(adrsTemp!presabs) + 1)
                                Else
                                    p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "     ") & _
                                    Spaces(Len(adrsTemp!presabs))
                                End If
                            Else
                                p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "     ") & _
                                Spaces(Len(adrsTemp!presabs))
                            End If
                       
           
                    End If
                   
                    
                ElseIf adrsTemp!Date <> dtTempDate Then
                    A_Str = A_Str & Spaces(0)
                    D_Str = D_Str & Spaces(0)
                    S_Str = S_Str & Spaces(0)
                    'comment by
                    L_Str = L_Str & Spaces(0)
                    E_str = E_str & Spaces(0)
                    ''
                    
                         
                        If typOptIdx.bytMon = 0 Or typOptIdx.bytMon = 5 Or typOptIdx.bytMon = 10 Or typOptIdx.bytMon = 11 Then
                            W_Str = W_Str & Spaces(0)
                            O_Str = O_Str & Spaces(0)
                            p_str = p_str & Spaces(0)
                        End If
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        If Trim(A_Str) = "" And Trim(D_Str) = "" And Trim(L_Str) = "" _
        And Trim(E_str) = "" And Trim(W_Str) = "" And Trim(O_Str) = "" _
        And Trim(p_str) = "" And Trim(S_Str) = "" Then
        Else
            ConMain.Execute "insert into " & strRepFile & "(Empcode," & strKDate & ",arrstr," & _
            "depstr,latestr,earlstr,workstr,otstr,presabsstr,shfstr,sumlate,sumearly,sumwork,sumOT) " & _
            " values(" & "'" & STRECODE & "'" & "," & "'" & strDateS & "'" & "," & "'" & A_Str & "'" & ",'" & _
            D_Str & "'" & ",'" & L_Str & "'" & "," & "'" & E_str & "'" & "," & "'" & W_Str & "'" & "," & _
            "'" & O_Str & "'" & ",'" & p_str & "'" & "," & "'" & S_Str & "'" & "," & lsum & _
            "," & esum & "," & wsum & "," & osum & ")"
        End If
    Loop
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbExclamation
    monPerfOt = False
End If
Exit Function
ERR_P:
    ShowError ("Monthly Performance / OT :: Reports")
    monPerfOt = False
    'Resume Next
End Function

'this function add by
Public Function GetHr2Days(mInput As Single) As Single
    Dim Temp As Single
    Temp = hrs2Dec(mInput)
    Temp = Temp / 8
    'temp = dec2Hrs(temp)
    GetHr2Days = Temp
End Function
Public Function monTotLtEr() As Boolean
On Error GoTo ERR_P
monTotLtEr = True
Dim formVal As Single, intCnt As Integer
Dim strGM As String, STRECODE As String, strTrnFile As String
'strTrnFile = ""
Dim rs As New ADODB.Recordset

strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")
If adrsEmp.State = 1 Then adrsEmp.Close

    adrsEmp.Open "select empmst.Empcode,Name,empmst.cat from " & rpTables & " where " & _
    "Empcode = Empcode " & strSql & " order by Empcode", ConMain, adOpenStatic

If Not (adrsEmp.BOF And adrsEmp.EOF) Then
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp!Empcode
        'calculate total late
        formVal = 0
        If rs.State = 1 Then rs.Close
        Select Case typOptIdx.bytMon
            Case 15
                If adrsTemp.State = 1 Then adrsTemp.Close
 
                
                rs.Open "select count(earlhrs)  from " & strTrnFile & " where Empcode=" & _
                "'" & STRECODE & "' and earlhrs>0 and dflg='' AND presabs='P P ' ", ConMain, adOpenStatic
                
                    adrsTemp.Open "select count(latehrs)  from " & strTrnFile & " where Empcode=" & _
                    "'" & STRECODE & "' and latehrs>0 and dflg='' AND presabs='P P ' ", ConMain, adOpenStatic
          
                If adrsTemp(0) = 0 And rs(0) = 0 Then GoTo nuxt
                If Not (adrsTemp.BOF And adrsTemp.EOF) Then intCnt = adrsTemp(0) + IIf(GetFlagStatus("CombineLateEarly"), rs(0), 0)
                If AdrsCat.State = 1 Then AdrsCat.Close
                AdrsCat.Open "select ltinmnth,erinmnth,everlet,letcut from catdesc where cat=" & _
                "'" & adrsEmp!cat & "'" & " and laterule = 'Y'", ConMain, adOpenStatic
                If Not (AdrsCat.BOF And AdrsCat.EOF) Then
                    If intCnt > AdrsCat!ltinmnth And AdrsCat!everlet > 0 Then
                       If Not UCase(GetFlagStatus("CombineLateEarly")) Then  'Fire 2019
                            formVal = CInt((intCnt - AdrsCat!ltinmnth) \ AdrsCat!everlet) * AdrsCat!letcut  ' 15-07
                        Else
                            formVal = CInt((intCnt - AdrsCat!ltinmnth - AdrsCat!erinmnth) \ AdrsCat!everlet) * AdrsCat!letcut
                        End If
                    End If
                End If
            Case 16
                If adrsTemp.State = 1 Then adrsTemp.Close
                
                rs.Open "select count(latehrs)  from " & strTrnFile & " where Empcode=" & _
                    "'" & STRECODE & "' and latehrs>0 and dflg=''  AND presabs='P P ' ", ConMain, adOpenStatic
                    
                adrsTemp.Open "select count(earlhrs)  from " & strTrnFile & " where Empcode=" & _
                "'" & STRECODE & "' and earlhrs>0 and dflg='' AND presabs='P P ' ", ConMain, adOpenStatic
                If adrsTemp(0) = 0 And rs(0) = 0 Then GoTo nuxt
                If Not (adrsTemp.BOF And adrsTemp.EOF) Then intCnt = adrsTemp(0) + IIf(GetFlagStatus("CombineLateEarly"), rs(0), 0)
                If AdrsCat.State = 1 Then AdrsCat.Close
                AdrsCat.Open "select ltinmnth,erinmnth,evererl,erlcut,everlet,letcut from catdesc where cat=" & _
                "'" & adrsEmp!cat & "'" & " and laterule = 'Y'", ConMain, adOpenStatic
                If Not (AdrsCat.BOF And AdrsCat.EOF) Then
                    If intCnt > AdrsCat!erinmnth And AdrsCat!evererl > 0 Then
                       If Not UCase(GetFlagStatus("CombineLateEarly")) Then  'Fire 2019
                            formVal = CInt((intCnt - AdrsCat!erinmnth) \ AdrsCat!evererl) * AdrsCat!erlCut  ' 15-07
                        Else
                            formVal = CInt((intCnt - AdrsCat!ltinmnth - AdrsCat!erinmnth) \ AdrsCat!everlet) * AdrsCat!letcut
                        End If
                    End If
                End If
        End Select
        ConMain.Execute "insert into " & strRepFile & "(Empcode,daysded,lvval) " & _
        "values (" & "'" & STRECODE & "'" & ",'" & IIf(Not IsNull(formVal) And formVal > 0, _
        Format(formVal, "0.00"), 0) & "','" & IIf(Not IsNull(intCnt) And intCnt > 0, _
        Format(intCnt, "0.00"), 0) & "')"
nuxt:
        adrsEmp.MoveNext
    Loop
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "Select * From " & strRepFile & " Where 1=2", ConMain, adOpenDynamic, adLockOptimistic
End If
Exit Function
ERR_P:
    ShowError ("Monthly Total Late Early :: Reports")
    monTotLtEr = False
    'Resume Next
End Function

Public Function monTotLtErEntire() As Boolean
On Error GoTo ERR_P                                         ' 26-11
monTotLtErEntire = True                                     'For Calculate Late and Early Days Deducted Entirely
Dim formVal As Single, intCnt As Integer
Dim strGM As String, STRECODE As String, ListDate As String

'strMon_Trn = "Lvtrn" & Right(GetTrnYear(DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))), 2)
If CByte(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then     ' 27-01
    strMon_Trn = "lvtrn" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
Else
    strMon_Trn = "lvtrn" & Right(typRep.strMonYear, 2)
End If

ListDate = strDTEnc & Format(DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "1")), "mm/dd/yy") & strDTEnc
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select empmst.Empcode,Name,empmst.cat from " & rpTables & " where " & _
"Empcode = Empcode " & strSql & " order by Empcode", ConMain, adOpenStatic
If Not (adrsEmp.BOF And adrsEmp.EOF) Then
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp!Empcode
        'calculate total late
        formVal = 0
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select lt_no  from " & strMon_Trn & " where lst_date = " & ListDate & " and Empcode = '" & STRECODE & "' " _
        , ConMain, adOpenStatic
        If adrsTemp.RecordCount <> 0 Then
            If Not (adrsTemp.BOF And adrsTemp.EOF) Then intCnt = adrsTemp(0)
            If AdrsCat.State = 1 Then AdrsCat.Close
            AdrsCat.Open "select ltinmnth,everlet,letcut from catdesc where cat=" & _
            "'" & adrsEmp!cat & "'" & " and laterule = 'Y'", ConMain, adOpenStatic
            If Not (AdrsCat.BOF And AdrsCat.EOF) Then
                If intCnt > AdrsCat!ltinmnth And AdrsCat!everlet > 0 Then
        '                    formVal = CInt((intCnt - AdrsCat!ltinmnth) / AdrsCat!everlet) * AdrsCat!letcut
                    formVal = CInt(Left(CStr((intCnt - AdrsCat!ltinmnth) / AdrsCat!everlet), 2)) * AdrsCat!letcut
                End If
            End If
        End If
        adrsTemp.Close
        adrsTemp.Open "select erl_no  from " & strMon_Trn & " where lst_date = " & ListDate & " and Empcode = '" & STRECODE & "' " _
        , ConMain, adOpenStatic
        If adrsTemp.RecordCount = 0 Then GoTo nuxt
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then intCnt = adrsTemp(0)
        If AdrsCat.State = 1 Then AdrsCat.Close
        AdrsCat.Open "select erinmnth,evererl,erlcut from catdesc where cat=" & _
        "'" & adrsEmp!cat & "'" & " and Earlrule = 'Y'", ConMain, adOpenStatic
        If Not (AdrsCat.BOF And AdrsCat.EOF) Then
            If intCnt > AdrsCat!erinmnth And AdrsCat!evererl > 0 Then
                formVal = formVal + CInt(Left(CStr((intCnt - AdrsCat!erinmnth) / AdrsCat!evererl), 2)) * AdrsCat!erlCut
            End If
        End If
        ConMain.Execute "insert into " & TempFile & "(Empcode,daysded,lvval) " & _
        "values (" & "'" & STRECODE & "'" & ",'" & IIf(Not IsNull(formVal) And formVal > 0, _
        Format(formVal, "0.00"), "0.00") & "','" & IIf(Not IsNull(intCnt) And intCnt > 0, _
        Format(intCnt, "0.00"), 0) & "')"
nuxt:
        adrsEmp.MoveNext
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Monthly Total Late and Early :: Reports")
    monTotLtErEntire = False
End Function

Private Function SumLeave(ByVal bytTRCD As Byte, ByVal STRECODE As String, _
ByVal strLdt As String, ByVal strLeaveC As String) As Single
On Error GoTo ERR_P
Dim strFileName As String

SumLeave = 0

'''''''''''
If typDlyLvBal.bytDtOpt = 1 Then
    If Val(pVStar.Yearstart) > MonthNumber(typDlyLvBal.strMnth) Then
        strFileName = "lvinfo" & Right(CStr(CInt(typDlyLvBal.strYr) - 1), 2)
    Else
        strFileName = "lvinfo" & Right(typDlyLvBal.strYr, 2)
    End If
Else
    If Val(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then
        strFileName = "lvinfo" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
    Else
        strFileName = "lvinfo" & Right(typRep.strMonYear, 2)
    End If
End If
If Not FindTable(strFileName) Then Exit Function
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open " select sum(days) from " & strFileName & " where lcode = '" & strLeaveC & _
"' and trcd= " & bytTRCD & " and fromdate <= " & strDTEnc & Format(DateCompDate(strLdt), "dd/mmm/yy") & strDTEnc & _
" and Empcode = '" & STRECODE & "'", ConMain, adOpenKeyset
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    SumLeave = IIf(IsNull(adrsTemp(0)), 0, adrsTemp(0))
End If
Exit Function
ERR_P:
    ShowError ("Sum Leave  :: Reports")
End Function

Private Function SumAvail(ByVal STRECODE As String, ByVal strFdate As String, _
ByVal strLdate As String, ByVal strLvCode As String) As Single
On Error GoTo ERR_P
Dim strFileName As String
SumAvail = 0
''''
If typDlyLvBal.bytDtOpt = 1 Then
    If Val(pVStar.Yearstart) > MonthNumber(typDlyLvBal.strMnth) Then
        strFileName = "lvtrn" & Right(CStr(CInt(typDlyLvBal.strYr) - 1), 2)
    Else
        strFileName = "lvtrn" & Right(typDlyLvBal.strYr, 2)
    End If
Else
    If Val(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then
        strFileName = "lvtrn" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
    Else
        strFileName = "lvtrn" & Right(typRep.strMonYear, 2)
    End If
End If

If Not FindTable(strFileName) Then Exit Function    '
'SumAvail = SumAvail + SumLeave(4, strECode, strFdate, strLvCode)
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select sum(" & strLvCode & ") from " & strFileName & " where lst_date <= " & _
strDTEnc & Format(DateCompDate(strLdate), "dd/mmm/yy") & strDTEnc & " and Empcode = '" & STRECODE & "' ", ConMain, adOpenStatic
 If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    SumAvail = SumAvail + IIf(IsNull(adrsTemp(0)), 0, adrsTemp(0))
End If
Exit Function
ERR_P:
    ShowError ("Sum Avail :: Reports")
End Function
'***
Private Function SumDlyAvail(ByVal STRECODE As String, ByVal strLdate As String, ByVal strLvCode As String) As Single
On Error GoTo ERR_P
Dim strFileName As String

SumDlyAvail = 0
strFileName = Left(typDlyLvBal.strMnth, 3) & Right(typDlyLvBal.strYr, 2) & "trn"

If Not FindTable(strFileName) Then Exit Function
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select presabs from " & strFileName & " where empcode='" & STRECODE & "' and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(strLdate), "dd/mmm/yy") & strDTEnc & " order by " & strKDate, ConMain, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    While Not (adrsTemp.EOF)
        If Left(adrsTemp(0), 2) = strLvCode Then
            SumDlyAvail = SumDlyAvail + 0.5
            If SubLeaveFlag = 1 And Left(adrsTemp(0), 2) = "CM" Then SumDlyAvail = SumDlyAvail + 0.5   ' 07-11
        End If
        If Right(adrsTemp(0), 2) = strLvCode Then
            SumDlyAvail = SumDlyAvail + 0.5
            If SubLeaveFlag = 1 And Right(adrsTemp(0), 2) = "CM" Then SumDlyAvail = SumDlyAvail + 0.5   ' 07-11
        End If
        adrsTemp.MoveNext
    Wend
End If
Exit Function
ERR_P:
    ShowError ("SumDlyAvail :: Reports")
End Function


Public Function monLeaveBal() As Boolean
On Error GoTo ERR_P
monLeaveBal = True

Dim strLvInfo As String, strLvTrn As String
Dim sngOpen As Single, sngCredit As Single, sngEnc As Single, sngAvl As Single, sngDlyAvl As Single
Dim sngEar As Single, sngLate As Single, TotalLeave As Single
Dim strLeave As String, lvval As String, lvstr As String
Dim STRECODE As String, strFdate As String, strLdate As String
Dim dt As String
                    
If SubLeaveFlag = 1 Then ' 15-10
    Dim LvCodeArr() As String, LvValArr() As String, tmpLvCode As String, tmpLvVal As String
    Dim LvOrder As Variant
    Dim k As Integer, j As Integer, m As Integer, Flag As Integer, ElFlag As Integer, cnt As Integer
    Dim ELTot As Single, SLTot As Single
    LvOrder = Array("OD", "CL", "CO", "CM", "HP", "SL", "EN", "NE", "EL", "LW")
End If

If typDlyLvBal.bytMnthOpt = 1 Then
    strFdate = DateCompStr(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "F"))
    strLdate = DateCompStr(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L"))
End If

If typDlyLvBal.bytDtOpt = 1 Then
    dt = Day(typDlyLvBal.DailyDt)
End If

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select distinct Empcode,empmst.cat from " & rpTables & " where " & _
" Empcode=Empcode " & strSql & " order by Empcode ", ConMain, adOpenKeyset
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    If adrsLeave.State = 1 Then adrsLeave.Close
    adrsLeave.Open "select * from Leavbal", ConMain, adOpenKeyset
        For i = 1 To adrsLeave.Fields.Count - 1
            If SubLeaveFlag = 1 Then   ' 07-11
                If adrsLeave.Fields(i).name <> "HP" And adrsLeave.Fields(i).name <> "CM" Then
                    cnt = cnt + 1
                    lvstr = lvstr & adrsLeave.Fields(i).name & Spaces(Len(Trim(adrsLeave.Fields(i).name)) - 1)
                End If
            Else
                lvstr = lvstr & adrsLeave.Fields(i).name & Spaces(Len(Trim(adrsLeave.Fields(i).name)) - 1)
            End If
        Next
        If SubLeaveFlag = 1 Then
            ReDim LvCodeArr(cnt - 1): ReDim LvValArr(cnt - 1)
        End If
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp("Empcode")
            For i = 1 To adrsLeave.Fields.Count - 1
                If adrsDept1.State = 1 Then adrsDept1.Close
                adrsDept1.Open "select lvcode from leavdesc where cat = '" & adrsEmp("cat") & _
                "' and lvcode = '" & adrsLeave.Fields(i).name & "'", ConMain
                If Not (SubLeaveFlag = 1 And (adrsLeave.Fields(i).name = "EL")) Then        ' 07-11
                    If Not (adrsDept1.EOF And adrsDept1.BOF) Then
                         strLeave = adrsLeave.Fields(i).name
                        If typDlyLvBal.bytMnthOpt = 1 Then
                            If (SubLeaveFlag = 1 And (adrsLeave.Fields(i).name = "CM" Or adrsLeave.Fields(i).name = "HP")) = False Then   ' 07-11
                                sngOpen = sngOpen + SumLeave(1, STRECODE, strLdate, strLeave)
                                sngCredit = sngCredit + SumLeave(2, STRECODE, strLdate, strLeave)
                            End If
                            If (SubLeaveFlag = 1 And (adrsLeave.Fields(i).name = "SL")) = False Then   ' 07-11
                                sngEnc = sngEnc + SumLeave(3, STRECODE, strLdate, strLeave)
                                sngAvl = sngAvl + SumAvail(STRECODE, strFdate, strLdate, strLeave)
                                sngEar = sngEar + SumLeave(6, STRECODE, strLdate, strLeave)
                                sngLate = sngLate + SumLeave(7, STRECODE, strLdate, strLeave)
                            End If
                        ElseIf typDlyLvBal.bytDtOpt = 1 Then                                 ' changes for upto date leave balance]
                            If (SubLeaveFlag = 1 And (adrsLeave.Fields(i).name = "CM" Or adrsLeave.Fields(i).name = "HP")) = False Then   ' 07-11
                                sngOpen = sngOpen + SumLeave(1, STRECODE, typDlyLvBal.DailyDt, strLeave)
                                sngCredit = sngCredit + SumLeave(2, STRECODE, typDlyLvBal.DailyDt, strLeave)
                            End If
                            If (SubLeaveFlag = 1 And (adrsLeave.Fields(i).name = "SL")) = False Then   ' 07-11
                                sngEnc = sngEnc + SumLeave(3, STRECODE, typDlyLvBal.DailyDt, strLeave)
                                sngAvl = sngAvl + SumAvail(STRECODE, typDlyLvBal.typFdate, typDlyLvBal.typLdate, strLeave)
                                If Not (dt = Day(typDlyLvBal.typLdate)) Then
                                    sngDlyAvl = sngDlyAvl + SumDlyAvail(STRECODE, typDlyLvBal.DailyDt, strLeave)
                                End If
                                sngEar = sngEar + SumLeave(6, STRECODE, typDlyLvBal.DailyDt, strLeave)
                                sngLate = sngLate + SumLeave(7, STRECODE, typDlyLvBal.DailyDt, strLeave)
                            End If
                        End If
                    End If
                    If (SubLeaveFlag = 1 And (adrsLeave.Fields(i).name = "SL")) = True Then   ' 07-11
                        TotalLeave = (sngOpen + sngCredit) - SLTot: SLTot = 0
                    Else
                        TotalLeave = (sngOpen + sngCredit) - (sngEnc + sngAvl + sngDlyAvl + sngEar + sngLate)
                    End If
                    lvval = lvval & IIf(TotalLeave <> 0, Spaces(Len(CStr(Format(TotalLeave, "0.00"))) - 1) & CStr(Format(TotalLeave, "0.00")), Space(7))
                End If
                If SubLeaveFlag = 1 And (adrsLeave.Fields(i).name <> "CM" And adrsLeave.Fields(i).name <> "HP") Then   ' 07-11
                    For k = 0 To UBound(LvOrder) - 1
                        If adrsLeave.Fields(i).name = LvOrder(k) Then
                            LvCodeArr(j) = adrsLeave.Fields(i).name & Spaces(Len(Trim(adrsLeave.Fields(i).name)) - 1)
                            LvValArr(j) = IIf(TotalLeave <> 0, Spaces(Len(CStr(Format(TotalLeave, "0.00"))) - 1) & CStr(Format(TotalLeave, "0.00")), Space(7))
                            If adrsLeave.Fields(i).name = "EL" Then m = j
                            j = j + 1
                            Flag = 1: Exit For
                        End If
                    Next
                    If Flag = 0 Then
                        tmpLvCode = tmpLvCode & adrsLeave.Fields(i).name & Spaces(Len(Trim(adrsLeave.Fields(i).name)) - 1)
                        tmpLvVal = tmpLvVal & IIf(TotalLeave <> 0, Spaces(Len(CStr(Format(TotalLeave, "0.00"))) - 1) & CStr(Format(TotalLeave, "0.00")), Space(7))
                    Else
                        Flag = 0
                    End If
                Else
                    SLTot = SLTot + (sngEnc + sngAvl + sngDlyAvl + sngEar + sngLate)
                End If
                If SubLeaveFlag = 1 And (adrsLeave.Fields(i).name = "EN" Or adrsLeave.Fields(i).name = "NE") Then ELTot = ELTot + TotalLeave   ' 07-11
                 sngOpen = 0: sngCredit = 0: sngEnc = 0: sngAvl = 0: sngDlyAvl = 0: sngEar = 0: sngLate = 0
                TotalLeave = 0
            Next
            If Trim(lvval) = "" Or Trim(lvstr) = "" Then
            Else
                If SubLeaveFlag = 1 Then   ' 07-11
                    lvstr = "": lvval = ""
                    If FieldExists("Leavbal", "EL") Then
                        LvValArr(m) = IIf(ELTot <> 0, Spaces(Len(CStr(Format(ELTot, "0.00"))) - 1) & CStr(Format(ELTot, "0.00")), Space(7))
                        ELTot = 0
                    End If
                    For k = 0 To UBound(LvCodeArr)
                        lvstr = lvstr & LvCodeArr(k)
                        lvval = lvval & LvValArr(k)
                    Next
                    lvstr = lvstr & tmpLvCode: lvval = lvval & tmpLvVal
                End If
                ConMain.Execute "insert into " & strRepFile & "(Empcode,leavestr," & _
                "lvval)" & " values('" & STRECODE & "','" & lvstr & "','" & lvval & "')"
            End If
            lvval = "":
            If SubLeaveFlag = 1 Then   ' 07-11
                Erase LvCodeArr: Erase LvValArr: tmpLvCode = "": tmpLvVal = "": j = 0: ReDim LvCodeArr(cnt - 1): ReDim LvValArr(cnt - 1)
            End If
        adrsEmp.MoveNext
        If adrsEmp.EOF Then Exit Do
    Loop
    DateStr = lvstr
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbExclamation
    monLeaveBal = False
End If
Exit Function
ERR_P:
    ShowError ("Monthly Leave Balance Exact :: Reports")
    monLeaveBal = False
End Function

Public Function monLvBalCry() As Boolean
On Error GoTo ERR_P
monLvBalCry = True
Dim adrsLvCD As New ADODB.Recordset
Dim strLvInfo As String, strLvTrn As String
Dim sngOpen As Single, sngCredit As Single, sngEnc As Single, sngAvl As Single
Dim sngEar As Single, sngLate As Single, TotalLeave As Single
Dim strLeave As String, lvval() As String
Dim STRECODE As String, strFdate As String, strLdate As String
Dim valFld As String
'**************************************************************************************
If adrsLvCD.State = 1 Then adrsLvCD.Close
adrsLvCD.Open "select * from leavbal", ConMain, adOpenStatic

    For i = 1 To adrsLvCD.Fields.Count - 1
       StrLvCD(i) = "'" & adrsLvCD.Fields(i).name & "'"
      strAlv(i) = "{" & "cmd." & adrsLvCD.Fields(i).name & "}"
      Next i


strFdate = DateCompStr(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "F"))
strLdate = DateCompStr(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L"))

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select distinct Empcode,empmst.cat from " & rpTables & " where " & _
" Empcode=Empcode " & strSql & " order by Empcode ", ConMain, adOpenKeyset
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    If adrsLeave.State = 1 Then adrsLeave.Close
    adrsLeave.Open "select * from Leavbal", ConMain, adOpenKeyset
'        For i = 1 To adrsLeave.Fields.Count - 1
'            lvstr = lvstr & adrsLeave.Fields(i).Name & Spaces(Len(Trim(adrsLeave.Fields(i).Name)) - 1)
'        Next
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp("Empcode")
        ReDim lvval(adrsLeave.Fields.Count)
            For i = 1 To adrsLeave.Fields.Count - 1
                If adrsDept1.State = 1 Then adrsDept1.Close
                adrsDept1.Open "select lvcode from leavdesc where cat = '" & adrsEmp("cat") & _
                "' and lvcode = '" & adrsLeave.Fields(i).name & "'", ConMain
                If Not (adrsDept1.EOF And adrsDept1.BOF) Then
                    strLeave = adrsLeave.Fields(i).name
                    sngOpen = sngOpen + SumLeave(1, STRECODE, strLdate, strLeave)
                    sngCredit = sngCredit + SumLeave(2, STRECODE, strLdate, strLeave)
                    sngEnc = sngEnc + SumLeave(3, STRECODE, strLdate, strLeave)
                    sngAvl = sngAvl + SumAvail(STRECODE, strFdate, strLdate, strLeave)
                    sngEar = sngEar + SumLeave(6, STRECODE, strLdate, strLeave)
                    sngLate = sngLate + SumLeave(7, STRECODE, strLdate, strLeave)
                End If
                TotalLeave = (sngOpen + sngCredit) - (sngEnc + sngAvl + sngEar + sngLate)
                lvval(i - 1) = Format(TotalLeave, "0.00")
                                
                sngOpen = 0: sngCredit = 0: sngEnc = 0: sngAvl = 0: sngEar = 0: sngLate = 0
                TotalLeave = 0
            Next
            valFld = "'" & STRECODE & "'"
            For i = 1 To adrsLeave.Fields.Count - 1
            valFld = valFld & "," & lvval(i - 1)
            Next i
            ConMain.Execute "insert into " & strRepFile & " values(" & valFld & ")"
            
            
        adrsEmp.MoveNext
        If adrsEmp.EOF Then Exit Do
    Loop
    'DateStr = lvstr
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbExclamation
    monLvBalCry = False
End If
Exit Function
ERR_P:
    ShowError ("Monthly Leave Balance Exact :: Reports")
    Resume Next
    monLvBalCry = False
End Function

Public Function monALERep() As Boolean
On Error GoTo ERR_P
monALERep = True
Dim bytLatcnt As Byte, bytErlcnt As Byte
Dim sngLathrs As Single, sngErlhrs As Single, sngAbscnt As Single
Dim strFileName As String, STRECODE As String

strFileName = Left(typRep.strMonMth, 3) & Right(typRep.strMonYear, 2) & "trn"
bytLatcnt = 0: bytErlcnt = 0: sngLathrs = 0: sngErlhrs = 0: sngAbscnt = 0
If adrsTemp.State = 1 Then adrsTemp.Close

adrsTemp.Open "select distinct " & strFileName & ".Empcode,latehrs,earlhrs,presabs," & _
"" & strKDate & " from " & strFileName & "," & rpTables & " where " & strFileName & ".Empcode = empmst.Empcode and ((" & _
strFileName & ".latehrs > 0) or (" & strFileName & ".earlhrs > 0) or (" & strFileName & _
".presabs = '" & ReplicateVal(pVStar.AbsCode, 2) & "') OR (" & LeftStr(strFileName & _
".presabs") & " = '" & pVStar.AbsCode & "') OR (" & RightStr(strFileName & ".presabs") & " = '" & _
pVStar.AbsCode & "')) " & strSql & " order by " & strFileName & ".Empcode," & strFileName & "." & strKDate & ""

If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    Do While Not adrsTemp.EOF
        bytLatcnt = 0: bytErlcnt = 0: sngLathrs = 0: sngErlhrs = 0: sngAbscnt = 0
        STRECODE = adrsTemp!Empcode
        Do While STRECODE = adrsTemp!Empcode And Not adrsTemp.EOF
            If adrsTemp!latehrs > 0 Then
                sngLathrs = TimAdd(sngLathrs, adrsTemp!latehrs)
                bytLatcnt = bytLatcnt + 1
            End If
            If adrsTemp!earlhrs > 0 Then
                sngErlhrs = TimAdd(sngErlhrs, adrsTemp!earlhrs)
                bytErlcnt = bytErlcnt + 1
            End If
            If adrsTemp!presabs = (pVStar.AbsCode & pVStar.AbsCode) Then
                sngAbscnt = sngAbscnt + 1
            ElseIf (Left(adrsTemp!presabs, 2) = pVStar.AbsCode) Or (Right(adrsTemp!presabs, 2) = pVStar.AbsCode) Then
                sngAbscnt = sngAbscnt + 0.5
            End If
            adrsTemp.MoveNext
            If adrsTemp.EOF Then Exit Do
        Loop
        ConMain.Execute " insert into " & strRepFile & _
        "(Empcode,Absent,Lateno,LateHrs,EarlyNo,EarlyHrs) values(" & "'" & _
        STRECODE & "'" & "," & "'" & IIf(sngAbscnt > 0, Format(sngAbscnt, "0.00"), Empty) & _
        "'" & "," & "'" & IIf(bytLatcnt > 0, Format(bytLatcnt, "0.00"), Empty) & _
        "'" & "," & "'" & IIf(sngLathrs > 0, Format(sngLathrs, "0.00"), Empty) & _
        "'" & "," & "'" & IIf(bytErlcnt > 0, Format(bytErlcnt, "0.00"), Empty) & "'" & "," & _
        "'" & IIf(sngErlhrs > 0, Format(sngErlhrs, "0.00"), Empty) & "'" & ")"
    Loop
Else
    Call SetMSF1Cap(10)
    monALERep = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If
Exit Function
ERR_P:
    ShowError ("Monthly A/L/E Reports :: Reports")
    monALERep = False
End Function

Public Function monAttendanceAPPM() As Boolean
    On Error GoTo Err
    monAttendanceAPPM = True
    Dim strMonthFile As String
    Dim strLvTrnFile As String
    Dim adrsEmp As ADODB.Recordset
    Dim adrsTempForTrn As ADODB.Recordset
    Dim adrsTempForLv As ADODB.Recordset
    Dim strAbsent As String
    Dim strEmpCode As String
    Dim strLeaveCode() As String
    Dim sngLeave As Single
    Dim strPrevMonth As String
    Dim intPrevMonthNo As Integer
    
    Dim intPrevYear As Integer
    Dim strPrevMonthForQuery As String
    
    strMonthFile = MakeName(typRep.strMonMth, typRep.strMonYear, "Trn")
    intPrevMonthNo = MonthNumber(typRep.strMonMth) - 1
    
    If intPrevMonthNo = 0 Then
        strPrevMonth = MakeName("December", Val(typRep.strMonYear) - 1, "Trn")
        intPrevYear = Val(typRep.strMonYear) - 1
        strPrevMonthForQuery = "Dec"
    Else
        strPrevMonth = MakeName(MonthName(intPrevMonthNo), typRep.strMonYear, "Trn")
        intPrevYear = Val(typRep.strMonYear)
        strPrevMonthForQuery = MonthName(MonthNumber(typRep.strMonMth) - 1)
    End If
    
    strMon_Trn = strMonthFile
    
    If Not FindTable(strMonthFile) Then
        MsgBox "Transaction file not present for this month", vbInformation
        monAttendanceAPPM = False
        Exit Function
    End If
    
    If Not FindTable(strPrevMonth) Then
        MsgBox "Transaction file of previous month not present", vbInformation
        monAttendanceAPPM = False
        Exit Function
    End If
    
    strLvTrnFile = "Lvtrn" & Right(typRep.strMonYear, 2)
    strMon_Trn2 = strLvTrnFile
    If Not FindTable(strLvTrnFile) Then
        MsgBox "Leave Transaction file not present for this year", vbInformation
        monAttendanceAPPM = False
        Exit Function
    End If
    
    Set adrsEmp = OpenRecordSet("SELECT Empcode FROM " & _
     rpTables & " WHERE " & Right(strSql, Len(strSql) - 4) & " ORDER BY empmst.empcode")
    Dim strTemp As String
    strLeaveCode = Split(GetLeaveCode("+", , " lvcode NOT IN ('" & _
    pVStar.AbsCode & "','" & pVStar.PrsCode & "','" & _
    pVStar.WosCode & "','" & pVStar.HlsCode & "','OD')"), "+")
    For i = 0 To UBound(strLeaveCode) - 1
        strTemp = strTemp & "IIF(ISNULL(" & _
            strLeaveCode(i) & "),0," & strLeaveCode(i) & ")" & "+"
    Next
    strTemp = Left(strTemp, Len(strTemp) - 1)
    If Not (adrsEmp.EOF And adrsEmp.BOF) Then
        Do While Not adrsEmp.EOF
           strAbsent = ""
           strEmpCode = adrsEmp.Fields("Empcode")
           Set adrsTempForTrn = OpenRecordSet("SELECT " & strKDate & _
           " FROM " & strPrevMonth & " WHERE " & strPrevMonth & _
           "." & strKDate & ">=" & strDTEnc & "" & DateCompStr("26/" & strPrevMonthForQuery & _
           "/" & intPrevYear) & "" & strDTEnc & " AND empcode='" & _
           strEmpCode & "' AND (LEFT(" & strPrevMonth & ".presabs,2)='" & _
           pVStar.AbsCode & "' OR RIGHT(" & strPrevMonth & ".presabs,2)='" & _
           pVStar.AbsCode & "') UNION SELECT " & strKDate & _
           " FROM " & strMonthFile & " WHERE " & strMonthFile & _
           "." & strKDate & "<=" & strDTEnc & "" & DateCompStr("25/" & typRep.strMonMth & _
           "/" & typRep.strMonYear) & "" & strDTEnc & " AND empcode='" & _
           strEmpCode & "' AND (LEFT(" & strMonthFile & ".presabs,2)='" & _
           pVStar.AbsCode & "' OR RIGHT(" & strMonthFile & ".presabs,2)='" & _
           pVStar.AbsCode & "') ORDER BY " & strKDate & "")
                Do While Not adrsTempForTrn.EOF
                    strAbsent = strAbsent & Day(adrsTempForTrn.Fields(0)) & ","
                    adrsTempForTrn.MoveNext
                Loop
           Set adrsTempForLv = OpenRecordSet("SELECT " & strTemp & _
           " AS 'LEAVES' FROM " & strLvTrnFile & " WHERE empcode='" & strEmpCode & "' AND lst_date=" & _
           strDTEnc & "" & GetMonthEnd(typRep.strMonMth, typRep.strMonYear) & "" & _
           strDTEnc & "")
           If Not (adrsTempForLv.BOF And adrsTempForLv.EOF) Then
                sngLeave = FilterNull(adrsTempForLv.Fields(0))
           End If
           ConMain.Execute "INSERT INTO " & strRepFile & _
           "(Empcode,Leave,DateOfAbsent) VALUES('" & strEmpCode & "','" & _
           sngLeave & "','" & strAbsent & "')"
           adrsEmp.MoveNext
        Loop
    End If
    Exit Function
Err:
    ShowError ("monAttendanceAPPM")
    monAttendanceAPPM = False
    Set adrsTempForLv = Nothing
    Set adrsTempForTrn = Nothing
End Function

Public Function monShiftSch() As Boolean
On Error GoTo ERR_P
monShiftSch = True
Dim strShfFile As String, S_Str As String
Dim dtFDate As Date, dtLDate As Date, dtTempDate As Date

strShfFile = Left(typRep.strMonMth, 3) & Right(typRep.strMonYear, 2) & "Shf"
dtFDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "F"))
dtLDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L"))
dtTempDate = dtFDate
DateStr = ""
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select " & strShfFile & ".* from " & strShfFile & "," & rpTables & _
" where " & strShfFile & ".Empcode = empmst.Empcode " & strSql & " order by " & _
strShfFile & ".Empcode", ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    DateStr = monDateStr(Day(dtLDate)) 'ASSIGNING VALUE TO GLOBAL VARIABLE -TO USE IN DSR
    Do While Not (adrsTemp.EOF) And dtTempDate < dtLDate
        S_Str = ""
        Do While dtTempDate <= dtLDate And Not (adrsTemp.EOF)
            If Trim(adrsTemp("d" & Day(dtTempDate))) <> "" Then
                S_Str = S_Str & adrsTemp("D" & Day(dtTempDate)) & Spaces(Len(adrsTemp("D" & Day(dtTempDate))))
            Else
                S_Str = S_Str & Spaces(0)
            End If
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        If Trim(S_Str) <> "" Then
            ConMain.Execute "insert into " & strRepFile & _
            "(Empcode,shfstr) values ('" & adrsTemp!Empcode & "','" & S_Str & "')"
        End If
        adrsTemp.MoveNext
        dtTempDate = dtFDate
    Loop
Else
    Call SetMSF1Cap(10)
    monShiftSch = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If
Exit Function
ERR_P:
    ShowError ("Monthly Shift Schedule :: Reports")
    monShiftSch = False
End Function

Public Function monAtMuPA() As Boolean
On Error GoTo ERR_P
monAtMuPA = True
Dim paidStr, otHrs, Wrk, night
Dim LVselect As String, strLvVal As String, lvstr As String
Dim strGM As String, STRECODE As String, DTESTR As String
Dim strTrnFile As String, strlvfile As String, p_str As String
Dim dtFDate As Date, dtLDate As Date, dtTempDate As Date
strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")
paidStr = 0: otHrs = 0: Wrk = 0: night = 0
If Val(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then
    strlvfile = "lvtrn" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
Else
    strlvfile = "lvtrn" & Right(typRep.strMonYear, 2)
End If

If Not FindTable(strlvfile) Then
    monAtMuPA = False
    MsgBox NewCaptionTxt("M7005", adrsMod), vbInformation
    Exit Function
End If
dtFDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "F"))
dtLDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))

dtTempDate = dtFDate

LVselect = ""
DTESTR = ""

If typOptIdx.bytMon = 1 Then 'START BYTMON=1
    lvstr = ""
    If adrsLeave.State = 1 Then adrsLeave.Close
    adrsLeave.Open "select distinct(lvcode) from leavdesc", ConMain, adOpenStatic
    If Not (adrsLeave.BOF And adrsLeave.EOF) Then
        For i = 1 To adrsLeave.RecordCount
            lvstr = lvstr & adrsLeave(0) & Spaces(Len(adrsLeave(0)))
            If i <> adrsLeave.RecordCount Then
                LVselect = LVselect & adrsLeave(0) & ","
            Else
                LVselect = LVselect & adrsLeave(0)
            End If
            adrsLeave.MoveNext
        Next i
    End If
End If 'END BYTMON=1

Select Case typOptIdx.bytMon
    Case 1, 2 'attendance, muster
        strGM = "select empmst.Empcode,presabs," & strKDate & ",arrtim from " & strTrnFile & "," & _
        rpTables & " where " & strTrnFile & ".Empcode = empmst.Empcode " & strSql & _
        " order by " & strTrnFile & ".Empcode," & strTrnFile & "." & strKDate & ""
    Case 3 'Month present
        strGM = "select empmst.Empcode,presabs," & strKDate & ",arrtim from " & strTrnFile & "," & _
        rpTables & " where empmst.Empcode = " & strTrnFile & ".Empcode and (presabs=" & "'" & ReplicateVal(pVStar.PrsCode, 2) & "'" & _
        " or " & LeftStr("presabs") & "=" & "'" & pVStar.PrsCode & "'" & " or " & RightStr("presabs") & "=" & _
        "'" & pVStar.PrsCode & "'" & ") " & strSql & " order by " & strTrnFile & ".Empcode," & _
        strTrnFile & "." & strKDate & ""
    Case 4 'Month absent
        strGM = "select empmst.Empcode,presabs," & strKDate & ",arrtim from " & strTrnFile & "," & _
        rpTables & " where empmst.Empcode = " & strTrnFile & ".Empcode and (presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
        " or " & LeftStr("presabs") & "=" & "'" & pVStar.AbsCode & "'" & " or " & RightStr("presabs") & "=" & _
        "'" & pVStar.AbsCode & "'" & ") " & strSql & " order by " & strTrnFile & _
        ".Empcode," & strTrnFile & "." & strKDate & ""
End Select
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGM, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    DTESTR = monDateStr(Day(dtLDate))
    Do While Not (adrsTemp.EOF) 'And dtTempDate < ldt
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtFDate
        p_str = ""
        Do While dtTempDate <= dtLDate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    If (Left(adrsTemp!presabs, 2) = pVStar.WosCode Or Left(adrsTemp!presabs, _
                    2) = pVStar.HlsCode) And adrsTemp!arrtim > 0 Then
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, _
                        "") & "p" & Spaces(Len(adrsTemp!presabs) + 1)
                    Else
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, _
                        "") & Spaces(Len(adrsTemp!presabs))
                    End If
                ElseIf adrsTemp!Date <> dtTempDate Then
                   p_str = p_str & Spaces(0)
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        If typOptIdx.bytMon = 1 Then 'START BYTMON = 1
            If adrsLeave.State = 1 Then adrsLeave.Close
            adrsLeave.Open "select paiddays,ot_hrs,wrk_hrs,night," & LVselect & " from " & strlvfile & " where Empcode =" & _
            "'" & STRECODE & "' and lst_date = " & strDTEnc & DateCompStr(dtLDate) & _
            strDTEnc, ConMain, adOpenStatic
            If Not (adrsLeave.EOF And adrsLeave.BOF) Then
                For i = 4 To adrsLeave.Fields.Count - 1
                    If Not IsNull(adrsLeave(i)) Then
                        strLvVal = strLvVal & IIf(Not IsNull(adrsLeave(i)) And adrsLeave(i) > 0, _
                        Format(adrsLeave(i), "0.00") & Spaces(Len(Format(adrsLeave(i), "0.00"))), Spaces(0))
                    Else
                        strLvVal = strLvVal & Spaces(0)
                    End If
                Next i
                paidStr = IIf(Not IsNull(adrsLeave!PaidDays) And adrsLeave!PaidDays > 0, _
                    Format(adrsLeave!PaidDays, "0.00"), Empty)
                otHrs = IIf(Not IsNull(adrsLeave!ot_hrs) And adrsLeave!ot_hrs > 0, _
                    Format(adrsLeave!ot_hrs, "0.00"), Empty)
                Wrk = IIf(Not IsNull(adrsLeave!wrk_Hrs) And adrsLeave!wrk_Hrs > 0, _
                    Format(adrsLeave!wrk_Hrs, "0.00"), Empty)
                night = IIf(Not IsNull(adrsLeave!night) And adrsLeave!night > 0, _
                    Format(adrsLeave!night, "0.00"), Empty)
            End If
        End If 'END BYTMON = 1
        If Trim(lvstr) = "" And Trim(strLvVal) = "" And Trim(p_str) = "" Then
        Else
            ConMain.Execute "insert into " & strRepFile & " " & _
            "(Empcode,mndatestr,presabsstr,leavestr,pdaysstr,otstr,wrkstr,nightstr,lvval)" & _
            " values(" & "'" & STRECODE & "'" & "," & "'" & DTESTR & "'" & "," & "'" & _
            p_str & "'" & "," & "'" & lvstr & "'" & "," & "'" & paidStr & "'" & "," & "'" & _
            otHrs & "'" & "," & "'" & Wrk & "'" & "," & "'" & night & "'" & "," & "'" & strLvVal & "'" & ")"
        End If
        strLvVal = ""
        p_str = "": night = "": Wrk = "": otHrs = "": paidStr = ""
        If bytBackEnd = 2 Then Sleep (500)
    Loop
End If
DateStr = DTESTR
Exit Function
ERR_P:
    ShowError ("Monthly Attendance Muster :: Reports")
''    Resume Next
    monAtMuPA = False
End Function

'

Public Function pePerfCryst()
On Error GoTo ERR_P
pePerfCryst = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim ArrArray(31), DepArray(31), LateArray(31), EarlArray(31), WrkArray(31), _
    OTarray(31) As Single
Dim PresArray(31), ShfArray(31) As String
Dim DateArray(31) As Integer
Dim ArrFld As Variant, ArrFld1 As Variant
Dim arrstr As Variant, ArrStr1 As Variant
Dim tmpStr As String

Dim STRECODE As String, strTrnFile As String, strDateS As String
Dim valFld As String, ValFld1 As String
Dim i, j
Dim strFld As String, strFld1 As String


Dim strTrnFile2 As String



'ReDim Preserve strarr(31)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strTrnFile = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strTrnFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

If typOptIdx.bytPer = 0 Or typOptIdx.bytPer = 1 Or typOptIdx.bytPer = 7 Then
    tmpStr = ""
ElseIf typOptIdx.bytPer = 2 Then
    tmpStr = " and ovtim > 0 "
ElseIf typOptIdx.bytPer = 3 Then
    tmpStr = " and latehrs > 0 "
ElseIf typOptIdx.bytPer = 4 Then
    tmpStr = " and earlhrs > 0 "
End If

If strTrnFile = strTrnFile2 Then
    If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select empmst.Empcode," & strKDate & _
            ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & "presabs," & _
            strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & _
            " where " & " empmst.Empcode = " & strTrnFile & ".Empcode " & tmpStr & _
            " and " & strKDate & ">=" & strDTEnc & DateCompStr( _
            typRep.strPeriFr) & strDTEnc & " and " & strKDate & "<=" & strDTEnc & _
            DateCompStr( _
            typRep.strPeriTo) & strDTEnc & strSql & " order by empmst.Empcode," & _
            strKDate & "", ConMain, adOpenStatic
Else
    If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select empmst.Empcode," & strKDate & _
            ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & "presabs," & _
            strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & _
            " where " & " empmst.Empcode = " & strTrnFile & ".Empcode  " & tmpStr & _
            " and " & strKDate & ">=" & strDTEnc & DateCompStr( _
            typRep.strPeriFr) & strDTEnc & " and " & strKDate & "<=" & strDTEnc & _
            DateCompStr( _
            typRep.strPeriTo) & strDTEnc & strSql & " Union select empmst.Empcode," _
            & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
            "presabs," & strTrnFile2 & ".shift,OTConf from " & strTrnFile2 & "," & _
            rpTables & " where " & " empmst.Empcode = " & strTrnFile2 & ".Empcode  " & tmpStr & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & strSql & " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
End If
strFld = "Empcode"
strFld1 = strFld

If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
    strDateS = monDateStr(Day(dttodate)) 'Assigning date string
    Do While Not (adrsTemp.EOF) 'And dtTempDate < dtToDate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        i = 1
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
              If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    DateArray(i) = Day(dtTempDate)
                    ArrArray(i) = IIf(Not IsNull( _
                        adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Format(adrsTemp!arrtim, "0.00"), 0)
                                     
                    DepArray(i) = IIf(Not IsNull( _
                        adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Format(adrsTemp!deptim, "0.00"), 0)
                    
                    ShfArray(i) = IIf(Not IsNull(adrsTemp!Shift), _
                        adrsTemp!Shift, "")
                    
                    LateArray(i) = IIf(Not IsNull( _
                        adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                        Format(adrsTemp!latehrs, "0.00"), 0)

                    EarlArray(i) = IIf(Not IsNull( _
                        adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                        Format(adrsTemp!earlhrs, "0.00"), 0)
  
                    WrkArray(i) = IIf(Not IsNull( _
                        adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                        Format(adrsTemp!wrkHrs, "0.00"), 0)
                        
                      
                        If adrsTemp("OTConf") = "Y" Then
                            OTarray(i) = IIf(Not IsNull( _
                                adrsTemp!ovtim) And adrsTemp!ovtim > 0, _
                                Format(adrsTemp!ovtim, "0.00"), 0)
                        Else
                            OTarray(i) = 0
                        End If
                        
                        If (Left(adrsTemp!presabs, _
                            2) = pVStar.WosCode Or Left(adrsTemp!presabs, _
                            2) = pVStar.HlsCode) And adrsTemp!arrtim > 0 Then
                            PresArray(i) = IIf(Not IsNull(adrsTemp!presabs), _
                                adrsTemp!presabs, "") & "p"
                        Else
                            PresArray(i) = IIf(Not IsNull(adrsTemp!presabs), _
                                adrsTemp!presabs, "")

                        End If
                
                ElseIf adrsTemp!Date <> dtTempDate Then
                    DateArray(i) = Day(dtTempDate)
                    ArrArray(i) = 0
                    DepArray(i) = 0
                    ShfArray(i) = ""
                    LateArray(i) = 0
                    EarlArray(i) = 0
                    WrkArray(i) = 0
                    PresArray(i) = ""
                    OTarray(i) = 0
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
            i = i + 1
        Loop
        valFld = ""
        ValFld1 = ""
        'Erase DateArray
        valFld = "'" & STRECODE & "'"
        ValFld1 = valFld
        ' to generate value string
        
        ' Performance
        If typOptIdx.bytPer = 0 Or typOptIdx.bytPer = 1 Or typOptIdx.bytPer = 7 Then
             For j = 0 To 6
        ArrFld = Array(ArrArray(), DepArray(), LateArray(), EarlArray(), _
            WrkArray(), OTarray(), DateArray(), PresArray(), ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
'                If j = 6 Then
'                    valfld = valfld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, ArrFld(j)(i))
                If (ArrFld(6)(i)) <> 0 Then
                    valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), Null, _
                        ArrFld(j)(i))
                End If
            Next i
            
        Next j
        
              
        
        If typOptIdx.bytPer = 1 Or typOptIdx.bytPer = 7 Then
         For j = 7 To 8
         
        ArrFld1 = Array(ArrArray(), DepArray(), LateArray(), EarlArray(), _
            WrkArray(), OTarray(), ShfArray(), PresArray(), DateArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If (ArrFld1(6)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)) Or ArrFld1(j)(i) = 0, "", _
                        ArrFld1(j)(i)) & "'"
                End If
            Next i
            
           Next j
   
             
        Else
         For j = 7 To 8
        ArrFld1 = Array(ArrArray(), DepArray(), LateArray(), EarlArray(), _
            WrkArray(), OTarray(), DateArray(), PresArray(), ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If (ArrFld1(6)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)), "", _
                        ArrFld1(j)(i)) & "'"
                End If
            Next i
            
        Next j
        End If
        End If
        If typOptIdx.bytPer = 2 Then
        
        For j = 0 To 6
        ArrFld = Array(ArrArray(), DepArray(), LateArray(), EarlArray(), _
            WrkArray(), OTarray(), DateArray(), PresArray(), ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If j = 6 Then
                    If (ArrFld(6)(i)) <> 0 Then
                        valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, _
                            ArrFld(j)(i))
                    End If
                ElseIf (ArrFld(5)(i)) <> 0 Then
                    valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, _
                        ArrFld(j)(i))
                End If
            Next i

        Next j
        For j = 7 To 8
        ArrFld1 = Array(ArrArray(), DepArray(), LateArray(), EarlArray(), _
            WrkArray(), OTarray(), DateArray(), PresArray(), ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If (ArrFld1(5)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)), "", _
                        ArrFld1(j)(i)) & "'"
                End If
            Next i

        Next j
       End If
        ' Late Arrival
        If typOptIdx.bytPer = 3 Then
        For j = 0 To 3
        ArrFld = Array(ArrArray(), DepArray(), LateArray(), DateArray(), _
            ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If j = 3 Then
                    If (ArrFld(3)(i)) <> 0 Then
                        valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, _
                            ArrFld(j)(i))
                    End If
                ElseIf (ArrFld(2)(i)) <> 0 Then
                    valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, _
                        ArrFld(j)(i))
                End If
            Next i
            
        Next j
        For j = 4 To 4
        ArrFld1 = Array(ArrArray(), DepArray(), LateArray(), DateArray(), _
            ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If (ArrFld1(2)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)), "", _
                        ArrFld1(j)(i)) & "'"
                End If
            Next i
            
        Next j
        End If
        
        ' Early Arrival
        If typOptIdx.bytPer = 4 Then
        For j = 0 To 3
        ArrFld = Array(ArrArray(), DepArray(), EarlArray(), DateArray(), _
            ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If j = 3 Then
                    If (ArrFld(3)(i)) <> 0 Then
                        valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, _
                            ArrFld(j)(i))
                    End If
                ElseIf (ArrFld(2)(i)) <> 0 Then
                    valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, _
                        ArrFld(j)(i))
                End If
            Next i
            
        Next j
        For j = 4 To 4
        ArrFld1 = Array(ArrArray(), DepArray(), EarlArray(), DateArray(), _
            ShfArray())

            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If (ArrFld(2)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)), "", _
                        ArrFld1(j)(i)) & "'"
                End If
            Next i
            
        Next j
        End If

        'to generate field string
        If typOptIdx.bytPer = 3 Then
            arrstr = Array("Arr", "Dep", "Late", "dt")
            ArrStr1 = Array("shf")
             strFld = "Empcode"
            strFld1 = "Empcode"
        ElseIf typOptIdx.bytPer = 4 Then
            arrstr = Array("Arr", "Dep", "Earl", "dt")
            ArrStr1 = Array("shf")
             strFld = "Empcode"
            strFld1 = "Empcode"
        Else
            arrstr = Array("Arr", "Dep", "Late", "Earl", "Work", "OT", "Dt")
            ArrStr1 = Array("Rem", "shf")
            strFld = "Empcode"
            strFld1 = "Empcode"
            
            
        End If
       
        For j = 0 To UBound(arrstr)
            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If typOptIdx.bytPer = 3 Or typOptIdx.bytPer = 4 Then
                    If j = 3 Then
                        If (ArrFld(3)(i)) <> 0 Then
                            strFld = strFld & "," & arrstr(j) & (i)
                        End If
                    ElseIf (ArrFld(2)(i)) <> 0 Then
                        strFld = strFld & "," & arrstr(j) & (i)
                    End If
                ElseIf typOptIdx.bytPer = 2 Then
                    If j = 6 Then
                        If (ArrFld(6)(i)) <> 0 Then
                            strFld = strFld & "," & arrstr(j) & (i)
                        End If
                    ElseIf (ArrFld(5)(i)) <> 0 Then
                        strFld = strFld & "," & arrstr(j) & (i)
                    End If
                Else
                'for muster Reports'
                 
                    If (ArrFld(6)(i)) <> 0 Then
                        strFld = strFld & "," & arrstr(j) & (i)
                    End If
                End If
            Next i
        Next j
        
        
        'for another table
         For j = 0 To UBound(ArrStr1)
            For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
                If typOptIdx.bytPer = 3 Or typOptIdx.bytPer = 4 Then
                    If j = 3 Then
                        If (ArrFld(3)(i)) <> 0 Then
                            strFld1 = strFld1 & "," & ArrStr1(j) & (i)
                        End If
                    ElseIf (ArrFld(2)(i)) <> 0 Then
                        strFld1 = strFld1 & "," & ArrStr1(j) & (i)
                    End If
                ElseIf typOptIdx.bytPer = 2 Then
                    If j = 6 Then
                        If (ArrFld1(6)(i)) <> 0 Then
                            strFld1 = strFld1 & "," & ArrStr1(j) & (i)
                        End If
                    ElseIf (ArrFld(5)(i)) <> 0 Then
                        strFld1 = strFld1 & "," & ArrStr1(j) & (i)
                    End If
                Else
                    If (ArrFld1(6)(i)) <> 0 Then
                        strFld1 = strFld1 & "," & ArrStr1(j) & (i)
                    End If
                End If
            Next i
        Next j
     
        ': Data is been splitted and entered into two table
        
        ConMain.Execute "insert into " & strRepFile & "(" & _
            strFld & ") values(" & valFld & ")"
         ConMain.Execute "insert into " & strRepMfile & "(" & _
             strFld1 & ") values(" & ValFld1 & ")"
         Erase ArrArray: Erase DepArray: Erase LateArray: Erase OTarray: Erase DateArray: Erase WrkArray: Erase EarlArray
          
             strFld = "": valFld = "": strFld1 = "": ValFld1 = ""
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Periodic Performance Overtime :: Reports")
    pePerfCryst = False
'    Resume Next
End Function


Public Function pePerfOvt() As Boolean           'by
On Error GoTo ERR_P
pePerfOvt = True

Dim PVal As Single, AbsVal As Single, WOVal As Single, HLVal As Single, PaidDys As Single, PaidLv As Single
Dim TotLate As Single, DedLate As Single, TotEarl As Single, DedEarl As Single
Dim TempAbs As Single, WOCnt As Single, HLCnt As Single, PLCnt As Single, COCnt As Single, MLCnt As Single, LPCnt As Single
Dim OACnt As Single, OPcnt As Single, SPcnt As Single, WPcnt As Single, SLCnt As Single, CLCnt As Single
Dim strLeave As String, tempStr As String, CatCode As String
Dim PRCnt As Single, PTCnt As Single, ACnt As Single, PPCnt As Single, ODcnt As Single
Dim ELCnt As Single
Dim StatusArr()
Dim cnt As Integer, Arrcnt As Integer, k As Integer, j As Integer
Dim TempAdrs As New ADODB.Recordset

Dim lsum As Single, esum As Single, wsum As Single, osum As Single, SumOT As Single
Dim A_Str As String, D_Str As String, L_Str As String, E_str As String
Dim W_Str As String, O_Str As String, p_str As String, S_Str As String
Dim strGP As String, strOvt As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim DTESTR As String 'strECode As String
Dim strfile1 As String, strFile2 As String, strFileName1 As String

FixedLvCode = ""
dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)
strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
strFileName1 = "lvinfo" & Right(dtfromdate, 2)
strOvt = ""
If typOptIdx.bytPer = 2 Then strOvt = " and ovtim>0 and OTConf ='Y' "
DTESTR = ""
Do While dtfromdate <= dttodate
    If dtfromdate = dttodate Then
        DTESTR = DTESTR & Day(dtfromdate)
    ElseIf dtfromdate <> dttodate Then
        DTESTR = DTESTR & Day(dtfromdate) & Spaces(Len(Trim(str(Day(dtfromdate)))))
    End If
    dtfromdate = DateAdd("d", 1, dtfromdate)
Loop
    
Dim FixLvCode() As Single
Dim LvCode() As String
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "SELECT DISTINCT lvcode FROM Leavdesc " & " WHERE lvcode NOT IN ('" & _
pVStar.AbsCode & "','" & pVStar.PrsCode & "','" & pVStar.WosCode & "','" & pVStar.HlsCode & "')", ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    cnt = adrsTemp.RecordCount - 1
    ReDim LvCode(cnt)
    Do While Not adrsTemp.EOF
        LvCode(k) = adrsTemp.Fields(0)
        FixedLvCode = FixedLvCode & Space(1) & adrsTemp.Fields(0)
        k = k + 1
        adrsTemp.MoveNext
    Loop
End If
k = 0: Arrcnt = 0
If strfile1 = strFile2 Then
    cutFile = strfile1
    cutFile1 = strFile2
    If typOptIdx.bytPer = 39 Then
    strGP = "select " & strfile1 & ".Empcode,catdesc.cat," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs, " & _
        "ovtim,presabs," & strfile1 & ".shift,OTConf from " & strfile1 & "," & rpTables & " where " & _
        strfile1 & ".Empcode = empmst.Empcode and empmst.leavdate is null and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & strOvt & " " & strSql
    Else
    If typOptIdx.bytPer = 52 Then
    If Not FieldExists(strfile1, "LunchLtHrs") Then ConMain.Execute ParseQuery("ALTER TABLE " & strfile1 & " ADD LunchLtHrs real null")
    strGP = "select " & strfile1 & ".Empcode,catdesc.cat," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs, " & _
        "ovtim,presabs," & strfile1 & ".shift,OTConf,LunchLtHrs from " & strfile1 & "," & rpTables & " where " & _
        strfile1 & ".Empcode = empmst.Empcode and empmst.leavdate is null and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & strOvt & " " & strSql
    Else
    strGP = "select " & strfile1 & ".Empcode,catdesc.cat," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs, " & _
        "ovtim,presabs," & strfile1 & ".shift,OTConf from " & strfile1 & "," & rpTables & " where " & _
        strfile1 & ".Empcode = empmst.Empcode and empmst.leavdate is null and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & strOvt & " " & strSql
    End If
    End If
    strFromDt = typRep.strPeriFr
    strToDt = typRep.strPeriTo
Else
    cutFile = strfile1
    cutFile1 = strFile2
    If typOptIdx.bytPer = 39 Then
    strGP = "select " & strfile1 & ".Empcode,catdesc.cat," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs," & _
        "ovtim,presabs," & strfile1 & ".shift,OTConf from " & strfile1 & "," & rpTables & " where " & _
        strfile1 & ".Empcode = empmst.Empcode and empmst.leavdate is null and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & strOvt & strSql & " union select " & strFile2 & ".Empcode,catdesc.cat," & strKDate & "," & _
        "arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim,presabs," & strFile2 & ".shift,OTConf from " & _
        strFile2 & "," & rpTables & " where " & strFile2 & ".Empcode = empmst.Empcode " & _
        "and empmst.leavdate is null and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & strOvt & " " & strSql
    Else
    strGP = "select " & strfile1 & ".Empcode,catdesc.cat," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs," & _
        "ovtim,presabs," & strfile1 & ".shift,OTConf from " & strfile1 & "," & rpTables & " where " & _
        strfile1 & ".Empcode = empmst.Empcode and empmst.leavdate is null and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & strOvt & strSql & " union select " & strFile2 & ".Empcode,catdesc.cat," & strKDate & "," & _
        "arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim,presabs," & strFile2 & ".shift,OTConf from " & _
        strFile2 & "," & rpTables & " where " & strFile2 & ".Empcode = empmst.Empcode " & _
        "and empmst.leavdate is null and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & strOvt & " " & strSql
    End If
End If
'strgp1 = " select trcd,days,lcode from " & strFileName1 & " where "
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
            STRECODE = adrsTemp!Empcode: CatCode = adrsTemp!cat
            dtfromdate = dtFirstDate
            A_Str = ""
            D_Str = ""
            L_Str = ""
            E_str = ""
            W_Str = ""
            O_Str = ""
            p_str = ""
            S_Str = ""
            Arrcnt = 1: ReDim FixLvCode(cnt)
            Do While dtfromdate <= dttodate
                If adrsTemp.EOF Then Exit Do
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtfromdate Then
                        A_Str = A_Str & IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), Spaces(0))
                                                        
                        D_Str = D_Str & IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), Spaces(0))
                                                        
                        L_Str = L_Str & IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                        Spaces(Len(Format(adrsTemp!latehrs, "0.00"))) & Format(adrsTemp!latehrs, "0.00"), Spaces(0))
                                                        
                        E_str = E_str & IIf(Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                        Spaces(Len(Format(adrsTemp!earlhrs, "0.00"))) & Format(adrsTemp!earlhrs, "0.00"), Spaces(0))
                                                        
                        W_Str = W_Str & IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                        Spaces(Len(Format(adrsTemp!wrkHrs, "0.00"))) & Format(adrsTemp!wrkHrs, "0.00"), Spaces(0))
                                                        
                                                        
                        If typOptIdx.bytPer = 52 Then
                            O_Str = O_Str & IIf(Not IsNull(adrsTemp!LunchLtHrs) And adrsTemp!LunchLtHrs > 0, _
                             Spaces(Len(Format(adrsTemp!LunchLtHrs, "0.00"))) & Format(adrsTemp!LunchLtHrs, "0.00"), Spaces(0))
                            osum = TimAdd(IIf(IsNull(osum), 0, osum), IIf(IsNull(adrsTemp!LunchLtHrs), 0, adrsTemp!LunchLtHrs))
                        Else
                            If adrsTemp("OTConf") = "Y" Then
                                O_Str = O_Str & IIf(Not IsNull(adrsTemp!ovtim) And adrsTemp!ovtim > 0, _
                                Spaces(Len(Format(adrsTemp!ovtim, "0.00"))) & Format(adrsTemp!ovtim, "0.00"), Spaces(0))
                                osum = TimAdd(IIf(IsNull(osum), 0, osum), IIf(IsNull(adrsTemp!ovtim), 0, adrsTemp!ovtim))
                            Else
                                O_Str = O_Str & Spaces(0)
                            End If
                        End If

                        
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                                                        
                        '''''''
                        If Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0 Then TotLate = TotLate + 1
                        If Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0 Then TotEarl = TotEarl + 1
                        
                        ReDim Preserve StatusArr(1 To Arrcnt)
                        StatusArr(Arrcnt) = adrsTemp!presabs
                        Arrcnt = Arrcnt + 1
                        If Not IsNull(adrsTemp!presabs) Then
                        strLeave = Left(adrsTemp!presabs, 2)
                            Select Case Left(adrsTemp!presabs, 2)
                                Case pVStar.AbsCode
                                    AbsVal = AbsVal + 0.5
                                Case pVStar.PrsCode
                                    PVal = PVal + 0.5
                                Case pVStar.WosCode
                                    WOVal = WOVal + 0.5
                                Case pVStar.HlsCode
                                    HLVal = HLVal + 0.5
                                Case Else
                                    If TempAdrs.State = 1 Then TempAdrs.Close
                                    TempAdrs.Open "Select paid from leavdesc where lvcode ='" & strLeave & "' and cat = '" & adrsTemp!cat & "'", ConMain, adOpenStatic
                                    If Not (TempAdrs.EOF And TempAdrs.BOF) And TempAdrs!paid = "Y" Then
                                        PaidLv = PaidLv + 0.5
                                    End If
                                    For i = 0 To UBound(LvCode) - 1
                                        If strLeave = LvCode(i) Then
                                            FixLvCode(i) = FixLvCode(i) + 0.5
                                            Exit For
                                        End If
                                    Next
                            End Select
                            strLeave = Right(adrsTemp!presabs, 2)
                            Select Case Right(adrsTemp!presabs, 2)
                                Case pVStar.AbsCode
                                    AbsVal = AbsVal + 0.5
                                Case pVStar.PrsCode
                                    PVal = PVal + 0.5
                                Case pVStar.WosCode
                                    WOVal = WOVal + 0.5
                                Case pVStar.HlsCode
                                    HLVal = HLVal + 0.5
                                Case Else
                                    If TempAdrs.State = 1 Then TempAdrs.Close
                                    TempAdrs.Open "Select paid from leavdesc where lvcode ='" & strLeave & "' and cat = '" & adrsTemp!cat & "'", ConMain, adOpenStatic
                                    If Not (TempAdrs.EOF And TempAdrs.BOF) And TempAdrs!paid = "Y" Then
                                        PaidLv = PaidLv + 0.5
                                    End If
                                        For i = 0 To UBound(LvCode) - 1
                                            If strLeave = LvCode(i) Then
                                                FixLvCode(i) = FixLvCode(i) + 0.5
                                                Exit For
                                        End If
                                    Next
                            End Select
                        End If
                        '**********
                        S_Str = S_Str & IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "") & Spaces(Len(adrsTemp!Shift))
                        
                        lsum = TimAdd(IIf(IsNull(lsum), 0, lsum), IIf(IsNull(adrsTemp!latehrs) Or adrsTemp!latehrs <= 0, 0, adrsTemp!latehrs))
                        esum = TimAdd(IIf(IsNull(esum), 0, esum), IIf(IsNull(adrsTemp!earlhrs) Or adrsTemp!earlhrs <= 0, 0, adrsTemp!earlhrs))
                        wsum = TimAdd(IIf(IsNull(wsum), 0, wsum), IIf(IsNull(adrsTemp!wrkHrs), 0, adrsTemp!wrkHrs))
                        SumOT = TimAdd(IIf(IsNull(SumOT), 0, SumOT), IIf(IsNull(adrsTemp!ovtim), 0, adrsTemp!ovtim))
              
                    ElseIf adrsTemp!Date <> dtfromdate Then
                        A_Str = A_Str & Spaces(0)
                        D_Str = D_Str & Spaces(0)
                        L_Str = L_Str & Spaces(0)
                        E_str = E_str & Spaces(0)
                        W_Str = W_Str & Spaces(0)
                        O_Str = O_Str & Spaces(0)
                        p_str = p_str & Spaces(0)
                        S_Str = S_Str & Spaces(0)
                    End If
                Else
                    Exit Do
                End If
                If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
                dtfromdate = DateAdd("d", 1, dtfromdate)
            Loop 'END OF DATE LOOP


            PaidDys = PVal + WOVal + HLVal + PaidLv
            If Trim(A_Str) = "" And Trim(D_Str) = "" And Trim(L_Str) = "" And Trim(E_str) = "" _
            And Trim(W_Str) = "" And Trim(O_Str) = "" And Trim(p_str) = "" And Trim(S_Str) = "" Then
            Else
                strECode1 = STRECODE
'                tempStr = Spaces(Len(Format(PaidDys, "0.00"))) & Format(PaidDys, "0.00") & Spaces(Len(Format(PVal, "0.00"))) & Format(PVal, "0.00") & Spaces(Len(Format(AbsVal, "0.00"))) & Format(AbsVal, "0.00") & Spaces(Len(Format(WOVal, "0.00"))) & Format(WOVal, "0.00") & Spaces(Len(Format(HLVal, "0.00"))) & Format(HLVal, "0.00") & _
'                Spaces(Len(Format(TotLate, "0.00"))) & Format(TotLate, "0.00") & Spaces(Len(Format(TotEarl, "0.00"))) & Format(TotEarl, "0.00") & Spaces(Len(Format(LtCut, "0.00"))) & Format(LtCut, "0.00") & Spaces(Len(Format(earlcut, "0.00"))) & Format(earlcut, "0.00")
                                
                For i = 0 To UBound(FixLvCode)
                    tempStr = tempStr & Spaces(Len(Format(FixLvCode(i), "0.00"))) & Format(FixLvCode(i), "0.00")
                Next
                If typOptIdx.bytPer = 51 Or typOptIdx.bytPer = 52 Then
                If osum > 0 Then
                ConMain.Execute "insert into " & strRepFile & "" & _
                "(Empcode," & strKDate & ",arrstr,depstr,latestr,earlstr,workstr,otstr," & _
                "presabsstr,shfstr,sumlate,sumearly,sumwork, SumOT)  values('" & STRECODE & _
                "','" & DTESTR & "','" & A_Str & "','" & D_Str & "','" & L_Str & "','" & _
                E_str & "','" & W_Str & "','" & O_Str & "','" & p_str & "','" & S_Str & _
                "'," & Round(lsum, 2) & "," & osum & "," & wsum & ", " & SumOT & ")"
                End If
                Else
                ConMain.Execute "insert into " & strRepFile & "" & _
                "(Empcode," & strKDate & ",arrstr,depstr,latestr,earlstr,workstr,otstr," & _
                "presabsstr,shfstr,sumlate,sumearly,sumwork, SumOT)  values('" & STRECODE & _
                "','" & DTESTR & "','" & A_Str & "','" & D_Str & "','" & L_Str & "','" & _
                E_str & "','" & W_Str & "','" & O_Str & "','" & p_str & "','" & S_Str & _
                "'," & Round(lsum, 2) & "," & osum & "," & wsum & ", " & SumOT & ")"
                End If
            End If
            lsum = 0: esum = 0: wsum = 0: osum = 0: Erase StatusArr: PVal = 0: AbsVal = 0: WOVal = 0: HLVal = 0: PaidDys = 0: PaidLv = 0: tempStr = "": TotLate = 0: TotEarl = 0: Erase FixLvCode
            WOCnt = 0: HLCnt = 0: PLCnt = 0: COCnt = 0: PRCnt = 0: PTCnt = 0: ELCnt = 0: ACnt = 0: PPCnt = 0: ODcnt = 0: WPcnt = 0
            dtfromdate = dtFirstDate
        Loop 'END OF EMPLOYEE LOOP
    End If
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    pePerfOvt = False
End If 'adrsTemp.eof
Exit Function
ERR_P:
    ShowError ("Periodic Performance Overtime :: Reports")
    pePerfOvt = False
'    Resume Next
End Function

Public Function peMuster() As Boolean
On Error GoTo ERR_P
peMuster = True
Dim strfile1 As String, strFile2 As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim strGP As String, DTESTR As String, STRECODE As String
Dim p_str As String

dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

If strfile1 = strFile2 Then
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",presabs,arrtim from " & strfile1 & "," & _
    rpTables & " where " & strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & _
    strDTEnc & DateCompStr(dtfromdate) & strDTEnc & " and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(dttodate) & strDTEnc & " " & strSql
Else
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",presabs,arrtim from " & strfile1 & "," & _
    rpTables & " where " & strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & _
    strDTEnc & DateCompStr(dtfromdate) & strDTEnc & " " & strSql & " union select " & _
    strFile2 & ".Empcode," & strKDate & ",presabs,arrtim from " & strFile2 & "," & rpTables & _
    " where " & strFile2 & ".Empcode = empmst.Empcode and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(dttodate) & strDTEnc & " " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select

DTESTR = ""
dtFirstDate = dtfromdate
Do While dtFirstDate <= dttodate
    If dtFirstDate = dttodate Then
        DTESTR = DTESTR & Day(dtFirstDate)
    ElseIf dtFirstDate <> dttodate Then
        DTESTR = DTESTR & Day(dtFirstDate) & Spaces(Len(Trim(str(Day(dtFirstDate)))))
    End If
    dtFirstDate = DateAdd("d", 1, dtFirstDate)
Loop

If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    Do While Not (adrsTemp.EOF) 'And dtFirstDate <  dtToDate
        STRECODE = adrsTemp!Empcode
        dtFirstDate = dtfromdate
        p_str = ""
        Do While dtFirstDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtFirstDate Then
                    If (Left(adrsTemp!presabs, 2) = pVStar.WosCode Or Left(adrsTemp!presabs, 2) = pVStar.HlsCode) And _
                    adrsTemp!arrtim > 0 Then
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & "p" & _
                        Spaces(Len(Format(adrsTemp!presabs, "0.00")) + 1)
                    Else
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & _
                        Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                    End If
                ElseIf adrsTemp!Date <> dtFirstDate Then
                    p_str = p_str & Spaces(0)
                End If
            Else
                Exit Do
            End If
            If dtFirstDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtFirstDate = DateAdd("d", 1, dtFirstDate)
        Loop
        If Trim(p_str) = "" Then
        Else
            ConMain.Execute "insert into " & strRepFile & _
            "(Empcode," & strKDate & ",presabsstr)" & _
            " values(" & "'" & STRECODE & "'" & "," & "'" & DTESTR & "'" & "," & _
            "'" & p_str & "'" & ")"
        End If
        Loop
    DateStr = DTESTR   'ASSIGNING VALUE OF DATE STRING TO GLOBAL VARIABLE, TO USE IN DSR
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    peMuster = False
End If
Exit Function
ERR_P:
    ShowError ("Periodic Muster :: Reports")
    peMuster = False
End Function

Public Function peLateEarl() As Boolean
On Error GoTo ERR_P
peLateEarl = True
Dim A_Str As String, D_Str As String, L_Str As String, E_str As String
Dim W_Str As String, p_str As String, S_Str As String
Dim osum As Single, lsum As Single, esum As Single
Dim DTESTR As String, strGW As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim STRECODE As String, strLaEr As String
Dim strfile1 As String, strFile2 As String

dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

If typOptIdx.bytPer = 3 Then
    strLaEr = "latehrs"
Else
    strLaEr = "earlhrs"
End If
strGW = ""
If strfile1 = strFile2 Then
    strGW = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs," & _
    "presabs," & strfile1 & ".shift from " & strfile1 & "," & rpTables & " where " & strfile1 & _
    ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & DateCompStr(dtfromdate) & _
    strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(dttodate) & strDTEnc & " and " & _
    strLaEr & ">0 " & strSql
Else
    strGW = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs," & _
    "presabs," & strfile1 & ".shift from " & strfile1 & "," & rpTables & " where " & strfile1 & _
    ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & DateCompStr(dtfromdate) & _
    strDTEnc & " and " & strfile1 & "." & strLaEr & ">0 " & strSql & _
    " union select " & strFile2 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs," & _
    "presabs," & strFile2 & ".shift from " & strFile2 & ", " & rpTables & " where " & strFile2 & _
    ".Empcode = empmst.Empcode and " & strKDate & "<=" & strDTEnc & DateCompStr(dttodate) & _
    strDTEnc & " and " & strFile2 & "." & strLaEr & ">0 " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGW = strGW & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGW = strGW & " order by Empcode," & strKDate
End Select
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGW, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        dtFirstDate = dtfromdate
        DTESTR = ""
        Do While dtFirstDate <= dttodate
            If dtFirstDate = dttodate Then
                DTESTR = DTESTR & Day(dtFirstDate)
            ElseIf dtFirstDate <> dttodate Then
                DTESTR = DTESTR & Day(dtFirstDate) & Spaces(Len(Trim(str(Day(dtFirstDate)))))
            End If
            dtFirstDate = DateAdd("d", 1, dtFirstDate)
        Loop
        dtFirstDate = dtfromdate
        Do While Not (adrsTemp.EOF) And dtFirstDate <= dttodate
            STRECODE = adrsTemp!Empcode
            dtFirstDate = dtfromdate
            A_Str = "": D_Str = "": L_Str = "": E_str = ""
            p_str = "": S_Str = ""
            lsum = 0: esum = 0
            Do While dtFirstDate <= dttodate And Not (adrsTemp.EOF)
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtFirstDate Then
                        A_Str = A_Str & IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), Spaces(0))
                        
                        D_Str = D_Str & IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), Spaces(0))
                        
                        L_Str = L_Str & IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                        Spaces(Len(Format(adrsTemp!latehrs, "0.00"))) & Format(adrsTemp!latehrs, "0.00"), Spaces(0))
                        
                        E_str = E_str & IIf(Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                        Spaces(Len(Format(adrsTemp!earlhrs, "0.00"))) & Format(adrsTemp!earlhrs, "0.00"), Spaces(0))
                        
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(adrsTemp!presabs))
                        
                        S_Str = S_Str & IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "") & Spaces(Len(adrsTemp!Shift))
                        
                        lsum = TimAdd(IIf(IsNull(lsum), 0, lsum), IIf(IsNull(adrsTemp!latehrs), 0, adrsTemp!latehrs))
                        esum = TimAdd(IIf(IsNull(esum), 0, esum), IIf(IsNull(adrsTemp!earlhrs), 0, adrsTemp!earlhrs))
                    ElseIf adrsTemp!Date <> dtFirstDate Then
                        A_Str = A_Str & Spaces(0)
                        D_Str = D_Str & Spaces(0)
                        L_Str = L_Str & Spaces(0)
                        E_str = E_str & Spaces(0)
                        p_str = p_str & Spaces(0)
                        S_Str = S_Str & Spaces(0)
                    End If
                Else
                    Exit Do
                End If
                If dtFirstDate = adrsTemp!Date Then adrsTemp.MoveNext
                dtFirstDate = DateAdd("d", 1, dtFirstDate)
            Loop
            If Trim(A_Str) = "" And Trim(D_Str) = "" And Trim(L_Str) = "" And _
            Trim(E_str) = "" And Trim(p_str) = "" And Trim(S_Str) = "" Then
            Else
                ConMain.Execute "insert into " & strRepFile & "" & _
                "(Empcode," & strKDate & ",arrstr,depstr,latestr,earlstr,presabsstr,shfstr,sumlate," & _
                "sumearly)  values(" & "'" & STRECODE & "'" & "," & "'" & DTESTR & "'" & "," & _
                "'" & A_Str & "'" & ",'" & D_Str & "'" & "," & _
                "'" & L_Str & "'" & "," & "'" & E_str & "'" & "," & _
                "'" & p_str & "'" & "," & "'" & S_Str & "'" & _
                "," & lsum & "," & esum & ")"
            End If
            lsum = 0: esum = 0
            dtFirstDate = dtfromdate
        Loop
    End If
Else
    Call SetMSF1Cap(10)
    peLateEarl = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If
Exit Function
ERR_P:
    ShowError ("Periodic Late Early :: Reports")
    peLateEarl = False
End Function

Public Function peContAbs() As Boolean
On Error GoTo RepErr
peContAbs = True
Dim bytTempCnt As Byte
Dim dte As Date, strFileName1 As String, strFileName2 As String
Dim strP_Str As String, strTempdt As String, STRECODE As String
Dim dtFrDate As Date, dtLaDate As Date
Dim bytNoDay As Byte

dtFrDate = DateCompDate(typRep.strPeriFr)
dtLaDate = DateCompDate(typRep.strPeriTo)

bytNoDay = DateDiff("d", dtFrDate, dtLaDate) + 1

strFileName1 = MakeName(MonthName(Month(dtFrDate)), Year(dtFrDate), "trn")
strFileName2 = MakeName(MonthName(Month(dtLaDate)), Year(dtLaDate), "trn")
bytTempCnt = 0
Dim strGP As String
    If strFileName1 = strFileName2 Then
        strGP = "select " & strFileName1 & ".Empcode,presabs," & strKDate & " from " & strFileName1 & _
        "," & rpTables & " where " & strFileName1 & ".Empcode = empmst.Empcode and " & _
        "" & strKDate & " >=" & strDTEnc & DateCompStr(dtFrDate) & strDTEnc & " and " & strKDate & "<=" & _
        strDTEnc & DateCompStr(dtLaDate) & strDTEnc & " And ( presabs <> '" & _
        pVStar.PrsCode & pVStar.PrsCode & "' ) " & strSql
    Else
        strGP = "select  " & strFileName1 & ".Empcode,presabs," & strKDate & " from " & strFileName1 & _
        "," & rpTables & " where " & strFileName1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & _
        strDTEnc & DateCompStr(dtFrDate) & strDTEnc & " And presabs <> '" & pVStar.PrsCode & pVStar.PrsCode & _
        "' " & strSql & _
        " union  select " & strFileName2 & ".Empcode,presabs," & strKDate & " from  " & strFileName2 & _
        "," & rpTables & " where " & strFileName2 & ".Empcode = empmst.Empcode and " & strKDate & "<=" & _
        strDTEnc & DateCompStr(dtLaDate) & strDTEnc & " And  presabs <>'" & pVStar.PrsCode & pVStar.PrsCode & _
        "' " & strSql
    End If
    Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strFileName1 & ".Empcode," & strFileName1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
    End Select
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open strGP, ConMain, adOpenStatic
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
        adrsTemp.MoveFirst
        dte = dtFrDate
        DateStr = ""
        strTempdt = ""
        If Day(dtFrDate) >= Day(dtLaDate) Then
            strTempdt = DateCompDate(dtFrDate)
            Do While strTempdt <= dtLaDate
                DateStr = DateStr & Day(strTempdt) & Spaces(Len(Trim(str(Day(strTempdt)))))
                strTempdt = CStr(DateCompDate(strTempdt) + 1)
            Loop
        Else
            i = Day(dtFrDate)
            For i = i To Day(dtLaDate)
                DateStr = DateStr & i & Spaces(Len(Trim(CStr(i))))
            Next i
        End If
         Do While Not (adrsTemp.EOF) And dte <= dtLaDate
            STRECODE = adrsTemp!Empcode
            dte = dtFrDate
            strP_Str$ = ""
            Do While dte >= dtFrDate And dte <= dtLaDate And Not (adrsTemp.EOF)
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dte Then
                        If adrsTemp!presabs <> pVStar.AbsCode & pVStar.AbsCode And _
                        adrsTemp!presabs <> pVStar.WosCode & pVStar.WosCode And _
                        adrsTemp!presabs <> pVStar.HlsCode & pVStar.HlsCode Then
                            strP_Str = ""
                            Do
                                adrsTemp.MoveNext
                                If adrsTemp.EOF Then Exit Do
                            Loop Until adrsTemp!Empcode <> STRECODE
                            Exit Do
                        Else
                            strP_Str = strP_Str & IIf(Not IsNull(adrsTemp!presabs), _
                            adrsTemp!presabs, "") & Spaces(Len(adrsTemp!presabs))
                            bytTempCnt = bytTempCnt + 1
                        End If
                    ElseIf adrsTemp!Date <> dte Then
                        strP_Str = strP_Str & Spaces(0)
                    End If
                Else
                    Exit Do
                End If
                If dte = adrsTemp!Date Then
                    adrsTemp.MoveNext
                    'dte = CDate(strFrDate)
                    dte = DateAdd("d", 1, dte)
                Else
                    dte = DateAdd("d", 1, dte)
                End If
            Loop
        If strP_Str <> "" And bytTempCnt = bytNoDay Then
            ConMain.Execute "insert into " & strRepFile & "(Empcode," & _
            "presabsstr)  values(" & "'" & STRECODE & "'" & ",'" & strP_Str & "'" & ")"
        End If
        bytTempCnt = 0
        dte = dtFrDate
    Loop
'DateStr = "" ' NEVER UNCOMMENT OR USE THIS STATEMENT.THIS VALUE IS USED BY RELATED DSR
Else '' If No Records found
    Call SetMSF1Cap(10)
    peContAbs = False
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If
If adrsTemp.State = 1 Then adrsTemp.Close
Exit Function
RepErr:
    ShowError ("Periodic Continuos Absent :: Reports")
    peContAbs = False
''    Resume Next
End Function
Public Function ContAbsPer(AbsDays As Integer)           '
On Error GoTo ERR_P
ContAbsPer = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim PresArray(31) As String
Dim DateArray(31) As Integer
Dim ArrFld As Variant, arrstr As Variant
Dim bytCnt As Byte
Dim STRECODE As String, strTrnFile As String, strTrnFile2 As String, valFld As String, strFld As String
Dim i, j, cnt
Dim Total As Single
Dim FlgS As Boolean

dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)
strTrnFile = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strTrnFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

If strTrnFile = strTrnFile2 Then
     If adrsTemp.State = 1 Then adrsTemp.Close
     Select Case InVar.strSer
        Case 1, 2 'SQL-Server,MS Access
        If GetFlagStatus("PRESENT") And typOptIdx.bytPer = 5 Then
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strTrnFile & ".arrtim>0 and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & "  And presabs IN('P P ','HLHL','WOWO','P A ','A P ','ODOD','OD P','P OD') " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        Else
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & "  And (Left(presabs,2)='A ' or Right(presabs,2) = 'A ') " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        End If
        Case 3 'Oracle
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & "  And (Lpad(presabs,2) ='A ' or Rpad(presabs,2) ='A ') " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
    End Select
Else
    If adrsTemp.State = 1 Then adrsTemp.Close
    Select Case InVar.strSer
        Case 1, 2 'SQL-server,MS Access
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Left(presabs,2) ='A ' or Right(presabs,2) ='A ') " & strSql & _
        " Union select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile2 & ".shift,OTConf from " & strTrnFile2 & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile2 & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Left(presabs,2) ='A ' or Right(presabs,2) ='A ')" & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
     Case 3 ''Oracle
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Lpad(presabs,2)='A ' or Rpad(presabs,2) ='A ') " & strSql & _
        " Union select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile2 & ".shift,OTConf from " & strTrnFile2 & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile2 & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Lpad(presabs,2) ='A ' or Rpad(presabs,2) ='A ')" & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
End Select
End If
arrstr = Array("shf", "Rem")
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
    Do While Not (adrsTemp.EOF) 'And dtTempDate <= dttodate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        cnt = 0
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    DateArray(cnt) = Day(dtTempDate)
                    PresArray(cnt) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")
                    cnt = cnt + 1
                Else
                    DateArray(cnt) = 0
                    PresArray(cnt) = ""
                    cnt = cnt + 1
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        valFld = "'" & STRECODE & "'"
        
        Dim Temp As Integer, start As Integer, k As Integer
        strFld = "Empcode"
        bytCnt = 0
            start = 0: Temp = 0
            For i = 0 To DateDiff("d", dtfromdate, dttodate) + 1
                Temp = i
                If DateArray(i) = 0 Then
                    If bytCnt >= AbsDays Then
                       bytCnt = 0
                       
                    Else
                        For k = start To Temp
                            PresArray(k) = ""
                        Next k
                        bytCnt = 0
                    End If
                    start = Temp + 1
                 Else
                    bytCnt = bytCnt + 1
                End If
            Next i
            
     For j = 0 To UBound(arrstr)
        For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
            strFld = strFld & "," & arrstr(j) & i
        Next i
     Next j
     
    For i = 0 To DateDiff("d", dtfromdate, dttodate)
        If PresArray(i) <> "" Then
            FlgS = True
            Exit For
        End If
    Next i
    If FlgS = True Then
    dtTempDate = dtfromdate
     For j = 1 To 2
         For i = 0 To DateDiff("d", dtfromdate, dttodate)
             If j = 1 Then
                 valFld = valFld & ",'" & Day(dtTempDate) & "'"
                 dtTempDate = DateAdd("d", 1, dtTempDate)
             ElseIf j = 2 Then
                 valFld = valFld & ",'" & PresArray(i) & "'"
             End If
         Next i
     Next j
     ConMain.Execute " insert into " & strRepMfile & "(" & strFld & ") values(" & valFld & ")"
    End If
    valFld = "": Erase DateArray: Erase PresArray: FlgS = False
    If adrsTemp.EOF Then Exit Do
    adrsTemp.MoveNext
Loop
End If
Exit Function
ERR_P:
    ShowError ("Continuous Absent :: Reports")
    ContAbsPer = False
End Function

Public Function peSummary() As Boolean
On Error GoTo ERR_P
Dim strFileName As String
Dim sngpp As Single, sngAA As Single, sngWO As Single, sngHL As Single
Dim snglv As Single, sngOD As Single, sngTot As Single
Dim sngPPP As Single, sngAAP As Single, sngWOP As Single, sngHLP As Single
Dim sngLVP As Single, sngODP As Single, sngTotP As Single
Dim intDept As Integer, bytCnt As Integer, strDept As String, intStrength As Integer
Dim sngGTot As String
Dim intOnOT As Integer, sngOtHrs As Single

Dim strFileName1 As String, strFileName2 As String
Dim dtfromdate As Date, dttodate As Date

dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strFileName1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFileName2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

Dim strGP As String
If strFileName1 = strFileName2 Then
    strGP = "Select DISTINCT deptdesc.DEPT, DEPTDESC." & strKDesc & "," & strFileName1 & _
            ".PRESABS," & strFileName1 & ".EmpCode, " & strFileName1 & "." & strKDate & ", " & _
            strFileName1 & ".od_from," & strFileName1 & ".ovtim From " & rpTables & _
            "," & strFileName1 & " WHERE empmst.Empcode = " & strFileName1 & ".EMPCODE " & _
            " and " & strFileName1 & "." & strKDate & " between  " & strDTEnc & _
            DateCompStr(dtfromdate) & strDTEnc & " and " & strDTEnc & _
            DateCompStr(dttodate) & strDTEnc & " " & strSql
Else
    strGP = "Select DISTINCT deptdesc.DEPT, DEPTDESC." & strKDesc & "," & strFileName1 & _
            ".PRESABS," & strFileName1 & ".EmpCode, " & strFileName1 & "." & strKDate & ", " & _
            strFileName1 & ".od_from," & strFileName1 & ".ovtim From " & rpTables & "," & _
            strFileName1 & " WHERE empmst.Empcode = " & strFileName1 & ".EMPCODE and " & _
            strFileName1 & "." & strKDate & " >= " & strDTEnc & DateCompStr(dtfromdate) & _
            strDTEnc & " " & strSql & _
            " union Select DISTINCT deptdesc.DEPT, DEPTDESC." & strKDesc & "," & strFileName2 & _
            ".PRESABS," & strFileName2 & ".EmpCode, " & strFileName2 & "." & strKDate & ", " & _
            strFileName2 & ".od_from," & strFileName2 & ".ovtim From " & rpTables & "," & _
            strFileName2 & " WHERE empmst.Empcode = " & strFileName2 & ".EMPCODE and " & _
            strFileName2 & "." & strKDate & " <= " & strDTEnc & DateCompStr(dttodate) & _
            strDTEnc & " " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        ''strGP = strGP & " order by " & strFileName1 & ".Empcode," & strFileName1 & "." & strKDate & ""
        strGP = strGP & " ORDER BY DEPTDESC.DEPT," & strFileName1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " ORDER BY DEPT," & strKDate & ""
End Select
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic, adLockReadOnly

If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    TotLeave = 0: Totpresent = 0: totAbsent = 0
    TotWkOff = 0: TotRec = 0: TotLate = 0: TotOT = 0
    Do While Not adrsTemp.EOF
        intDept = adrsTemp("dept")
        strDept = adrsTemp("desc")
        intStrength = getSTR(intDept)
        Do While intDept = adrsTemp("dept")
            Select Case Left(adrsTemp("presabs"), 2)
                Case pVStar.PrsCode
                    sngpp = sngpp + 0.5
                Case pVStar.AbsCode
                    sngAA = sngAA + 0.5
                Case pVStar.WosCode
                    sngWO = sngWO + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case Else
                    snglv = snglv + 0.5
            End Select
            Select Case Right(adrsTemp("presabs"), 2)
                Case pVStar.PrsCode
                    sngpp = sngpp + 0.5
                Case pVStar.AbsCode
                    sngAA = sngAA + 0.5
                Case pVStar.WosCode
                    sngWO = sngWO + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case Else
                    snglv = snglv + 0.5
            End Select
            sngTot = sngTot + 1
            If adrsTemp("ovtim") > 0 Then intOnOT = intOnOT + 1
            sngOtHrs = TimAdd(sngOtHrs, adrsTemp("ovtim"))

            If adrsTemp("od_from") > 0 Then sngOD = sngOD + 1
            adrsTemp.MoveNext
            If adrsTemp.EOF = True Then Exit Do
        Loop
        bytCnt = bytCnt + 1
        If sngpp > 0 Then sngPPP = Format((sngpp * 100) / sngTot, "00.00")
        If sngAA > 0 Then sngAAP = Format((sngAA * 100) / sngTot, "00.00")
        If sngWO > 0 Then sngWOP = Format((sngWO * 100) / sngTot, "00.00")
        If sngHL > 0 Then sngHLP = Format((sngHL * 100) / sngTot, "00.00")
        If snglv > 0 Then sngLVP = Format((snglv * 100) / sngTot, "00.00")
        sngGTot = Round(sngPPP + sngAAP + sngWOP + sngHLP + sngLVP)

        ''For DSR
        Totpresent = Totpresent + sngpp     ''PP total
        totAbsent = totAbsent + sngAA       ''AA total
        TotWkOff = TotWkOff + sngWO         ''WO total
        TotOT = TotOT + sngHL               ''HL total
        TotLeave = TotLeave + snglv         ''LV total
        TotLate = TotLate + sngOD           ''OD total
        intTotOnOt = intTotOnOt + intOnOT
        sngTotOTHrs = TimAdd(sngTotOTHrs, sngOtHrs)

        ConMain.Execute "insert into " & strRepFile & " values(" & _
        bytCnt & ",'" & strDept & "'," & intStrength & "," & sngTot & "," & sngpp & _
        "," & sngPPP & "," & sngAA & "," & sngAAP & "," & sngWO & "," & sngWOP & _
        "," & sngHL & "," & sngHLP & "," & snglv & "," & sngLVP & "," & intOnOT & "," & _
        sngOtHrs & "," & sngOD & ",0," & sngTot & "," & sngGTot & ")"
        ''insert statment here

        sngpp = 0: sngAA = 0: sngWO = 0: sngHL = 0
        snglv = 0: sngOD = 0: sngTot = 0
        sngPPP = 0: sngAAP = 0: sngWOP = 0: sngHLP = 0
        sngLVP = 0: sngODP = 0: sngTotP = 0
        intOnOT = 0: sngOtHrs = 0
    Loop
Else
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    Exit Function
End If
peSummary = True
Exit Function
ERR_P:
    ShowError ("Periodic Summary :: Reports")
    peSummary = False
    ''Resume Next
End Function
''For Mauritius 11-07-2003
Public Function peMealAl() As Boolean
On Error GoTo ERR_P
Dim lsum As Single, esum As Single, wsum As Single, osum As Single
Dim A_Str As String, D_Str As String, L_Str As String, E_str As String
Dim W_Str As String, O_Str As String, p_str As String, S_Str As String
Dim strGP As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim DTESTR As String, STRECODE As String
Dim strfile1 As String, strFile2 As String
Dim sngFrom As Single, sngTo As Single

dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

sngFrom = InputBox("Please Enter FROM Time in 00.00 Format.")
sngTo = InputBox("Please Enter TO Time in 00.00 Format.")

If strfile1 = strFile2 Then
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,wrkhrs, " & _
    "presabs," & strfile1 & ".shift,Ot_auth,Et_hrs from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & _
    DateCompStr(typRep.strPeriFr) & strDTEnc & " and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(typRep.strPeriTo) & strDTEnc & " And Deptim Between " & sngFrom & " And " & _
    sngTo & " " & strSql
Else
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,wrkhrs," & _
    "presabs," & strfile1 & ".shift,Ot_auth,Et_hrs from " & strfile1 & "," & rpTables & _
    " where " & strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & _
    DateCompStr(typRep.strPeriFr) & strDTEnc & " And Deptim Between " & sngFrom & " And " & _
    sngTo & " " & strSql & _
    " union select " & strFile2 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,wrkhrs," & _
    "presabs," & strFile2 & ".shift,Ot_auth,Et_hrs from " & strFile2 & "," & rpTables & _
    " where " & strFile2 & ".Empcode = empmst.Empcode and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(typRep.strPeriTo) & strDTEnc & " And Deptim Between " & sngFrom & " And " & _
    sngTo & " " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
            STRECODE = adrsTemp!Empcode
            dtfromdate = dtFirstDate
            A_Str = ""
            D_Str = ""
            L_Str = ""
            E_str = ""
            W_Str = ""
            O_Str = ""
            p_str = ""
            S_Str = ""
            Do While dtfromdate <= dttodate
                If adrsTemp.EOF Then Exit Do
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtfromdate Then
                        If adrsTemp!OT_auth > 0 Or adrsTemp!ET_hrs > 0 Then
                        A_Str = A_Str & IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), Spaces(0))
                                                        
                        D_Str = D_Str & IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), Spaces(0))
                                                        
                        L_Str = L_Str & IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                        Spaces(Len(Format(adrsTemp!latehrs, "0.00"))) & Format(adrsTemp!latehrs, "0.00"), Spaces(0))
                        ''For Mauritius 11-07-2003
                        ''This is First Meal String
                        If (Not IsNull(adrsTemp!OT_auth) And adrsTemp!OT_auth > 0) Then
                            If Len(adrsTemp!OT_auth) = 6 Then
                                E_str = E_str & adrsTemp!OT_auth
                            Else
                                E_str = E_str & Spaces(Len(adrsTemp!OT_auth)) & adrsTemp!OT_auth
                            End If
                        Else
                            E_str = E_str & Spaces(0)
                        End If
                        ''This is Second Meal String
                        If (Not IsNull(adrsTemp!ET_hrs) And adrsTemp!ET_hrs > 0) Then
                            If Len(adrsTemp!ET_hrs) = 6 Then
                                O_Str = O_Str & adrsTemp!ET_hrs
                            Else
                                O_Str = O_Str & Spaces(Len(adrsTemp!ET_hrs)) & adrsTemp!ET_hrs
                            End If
                        Else
                            O_Str = O_Str & Spaces(0)
                        End If
                        ''
                        W_Str = W_Str & IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                        Spaces(Len(Format(adrsTemp!wrkHrs, "0.00"))) & Format(adrsTemp!wrkHrs, "0.00"), Spaces(0))
                        
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                                                        
                        S_Str = S_Str & IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "") & Spaces(Len(adrsTemp!Shift))
                        
                        esum = esum + adrsTemp!OT_auth
                        osum = osum + adrsTemp!ET_hrs
                        lsum = TimAdd(IIf(IsNull(lsum), 0, lsum), IIf(IsNull(adrsTemp!latehrs) Or adrsTemp!latehrs <= 0, 0, adrsTemp!latehrs))
                        wsum = TimAdd(IIf(IsNull(wsum), 0, wsum), IIf(IsNull(adrsTemp!wrkHrs), 0, adrsTemp!wrkHrs))
                        If dtfromdate = dttodate Then
                            DTESTR = DTESTR & Day(dtfromdate)
                        ElseIf dtfromdate <> dttodate Then
                            DTESTR = DTESTR & Day(dtfromdate) & Spaces(Len(Trim(str(Day(dtfromdate)))))
                        End If
                        End If
                    ElseIf adrsTemp!Date <> dtfromdate Then
                    End If
                Else
                    Exit Do
                End If
                If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
                dtfromdate = DateAdd("d", 1, dtfromdate)
            Loop 'END OF DATE LOOP
            If Trim(E_str) = "" And Trim(O_Str) = "" Then
            Else
                ConMain.Execute "insert into " & strRepFile & "(Empcode," & _
                strKDate & ",arrstr,depstr,latestr,earlstr,workstr,otstr,presabsstr,shfstr," & _
                "sumlate,sumearly,sumwork,sumextra)  values('" & STRECODE & "','" & DTESTR & _
                "','" & A_Str & "','" & D_Str & "','" & L_Str & "','" & E_str & "','" & _
                W_Str & "','" & O_Str & "','" & p_str & "','" & S_Str & "'," & _
                Round(lsum, 2) & "," & esum & "," & wsum & "," & osum & ")"
            End If
            lsum = 0: esum = 0: wsum = 0: osum = 0: DTESTR = ""
            dtfromdate = dtFirstDate
        Loop 'END OF EMPLOYEE LOOP
    End If
    peMealAl = True
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If 'adrsTemp.eof
Exit Function
ERR_P:
    ShowError ("Periodic Meal Allowance :: Reports")
End Function
''
''For Mauritius 12-07-2003
Public Function pe8Punches() As Boolean
On Error GoTo ERR_P
Dim lsum As Single, esum As Single, wsum As Single, osum As Single
Dim A_Str As String, D_Str As String, L_Str As String, E_str As String
Dim W_Str As String, O_Str As String, p_str As String, S_Str As String

Dim ActO_str As String, ActI_str As String, Tim5_str As String, Tim6_str As String
Dim Tim7_str As String, Tim8_str As String

Dim strGP As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim DTESTR As String, STRECODE As String
Dim strfile1 As String, strFile2 As String

dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

DTESTR = ""
Do While dtfromdate <= dttodate
    If dtfromdate = dttodate Then
        DTESTR = DTESTR & Day(dtfromdate)
    ElseIf dtfromdate <> dttodate Then
        DTESTR = DTESTR & Day(dtfromdate) & Spaces(Len(Trim(str(Day(dtfromdate)))))
    End If
    dtfromdate = DateAdd("d", 1, dtfromdate)
Loop
 
If strfile1 = strFile2 Then
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs," & _
    "wrkhrs,ovtim,presabs," & strfile1 & ".shift,OTConf,ACTRT_O, ACTRT_I, TIME5, TIME6, " & _
    "TIME7, TIME8 from " & strfile1 & "," & rpTables & " where " & strfile1 & _
    ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & _
    DateCompStr(typRep.strPeriFr) & strDTEnc & " and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(typRep.strPeriTo) & strDTEnc & " " & strSql
Else
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs," & _
    "wrkhrs,ovtim,presabs," & strfile1 & ".shift,OTConf,ACTRT_O, ACTRT_I, TIME5, TIME6," & _
    " TIME7, TIME8 from " & strfile1 & "," & rpTables & " where " & strfile1 & _
    ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & _
    DateCompStr(typRep.strPeriFr) & strDTEnc & strSql & _
    " union select " & strFile2 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs," & _
    "wrkhrs,ovtim,presabs," & strFile2 & ".shift,OTConf,ACTRT_O, ACTRT_I, TIME5, TIME6," & _
    " TIME7, TIME8 from " & strFile2 & "," & rpTables & " where " & strFile2 & _
    ".Empcode = empmst.Empcode and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(typRep.strPeriTo) & strDTEnc & " " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
            STRECODE = adrsTemp!Empcode
            dtfromdate = dtFirstDate
            A_Str = ""
            D_Str = ""
            L_Str = ""
            E_str = ""
            W_Str = ""
            O_Str = ""
            p_str = ""
            S_Str = ""
            ActO_str = ""
            ActI_str = ""
            Tim5_str = ""
            Tim6_str = ""
            Tim7_str = ""
            Tim8_str = ""
            Do While dtfromdate <= dttodate
                If adrsTemp.EOF Then Exit Do
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtfromdate Then
                        A_Str = A_Str & IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), Spaces(0))
                                                        
                        D_Str = D_Str & IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), Spaces(0))
                                                        
                        L_Str = L_Str & IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                        Spaces(Len(Format(adrsTemp!latehrs, "0.00"))) & Format(adrsTemp!latehrs, "0.00"), Spaces(0))
                                                        
                        E_str = E_str & IIf(Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                        Spaces(Len(Format(adrsTemp!earlhrs, "0.00"))) & Format(adrsTemp!earlhrs, "0.00"), Spaces(0))
                                                        
                        W_Str = W_Str & IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                        Spaces(Len(Format(adrsTemp!wrkHrs, "0.00"))) & Format(adrsTemp!wrkHrs, "0.00"), Spaces(0))
                                                        
                        If adrsTemp("OTConf") = "Y" Then     ''if authorized OT then only Calculate and show
                            O_Str = O_Str & IIf(Not IsNull(adrsTemp!ovtim) And adrsTemp!ovtim > 0, _
                            Spaces(Len(Format(adrsTemp!ovtim, "0.00"))) & Format(adrsTemp!ovtim, "0.00"), Spaces(0))
                            osum = TimAdd(IIf(IsNull(osum), 0, osum), IIf(IsNull(adrsTemp!ovtim), 0, adrsTemp!ovtim))
                        Else
                            O_Str = O_Str & Spaces(0)
                        End If
                        
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                                                        
                        S_Str = S_Str & IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "") & Spaces(Len(adrsTemp!Shift))
                        
                        ActO_str = ActO_str & IIf(Not IsNull(adrsTemp!Actrt_O) And adrsTemp!Actrt_O > 0, _
                        Spaces(Len(Format(adrsTemp!Actrt_O, "0.00"))) & Format(adrsTemp!Actrt_O, "0.00"), Spaces(0))
                        
                        ActI_str = ActI_str & IIf(Not IsNull(adrsTemp!actrt_i) And adrsTemp!actrt_i > 0, _
                        Spaces(Len(Format(adrsTemp!actrt_i, "0.00"))) & Format(adrsTemp!actrt_i, "0.00"), Spaces(0))
                        
                        Tim5_str = Tim5_str & IIf(Not IsNull(adrsTemp!time5) And adrsTemp!time5 > 0, _
                        Spaces(Len(Format(adrsTemp!time5, "0.00"))) & Format(adrsTemp!time5, "0.00"), Spaces(0))
                        
                        Tim6_str = Tim6_str & IIf(Not IsNull(adrsTemp!time6) And adrsTemp!time6 > 0, _
                        Spaces(Len(Format(adrsTemp!time6, "0.00"))) & Format(adrsTemp!time6, "0.00"), Spaces(0))
                        
                        Tim7_str = Tim7_str & IIf(Not IsNull(adrsTemp!time7) And adrsTemp!time7 > 0, _
                        Spaces(Len(Format(adrsTemp!time7, "0.00"))) & Format(adrsTemp!time7, "0.00"), Spaces(0))
                        
                        Tim8_str = Tim8_str & IIf(Not IsNull(adrsTemp!time8) And adrsTemp!time8 > 0, _
                        Spaces(Len(Format(adrsTemp!time8, "0.00"))) & Format(adrsTemp!time8, "0.00"), Spaces(0))
                        
                        
                        lsum = TimAdd(IIf(IsNull(lsum), 0, lsum), IIf(IsNull(adrsTemp!latehrs) Or adrsTemp!latehrs <= 0, 0, adrsTemp!latehrs))
                        esum = TimAdd(IIf(IsNull(esum), 0, esum), IIf(IsNull(adrsTemp!earlhrs) Or adrsTemp!earlhrs <= 0, 0, adrsTemp!earlhrs))
                        wsum = TimAdd(IIf(IsNull(wsum), 0, wsum), IIf(IsNull(adrsTemp!wrkHrs), 0, adrsTemp!wrkHrs))
                        
                    ElseIf adrsTemp!Date <> dtfromdate Then
                        A_Str = A_Str & Spaces(0)
                        D_Str = D_Str & Spaces(0)
                        L_Str = L_Str & Spaces(0)
                        E_str = E_str & Spaces(0)
                        W_Str = W_Str & Spaces(0)
                        O_Str = O_Str & Spaces(0)
                        p_str = p_str & Spaces(0)
                        S_Str = S_Str & Spaces(0)
                        ActO_str = ActO_str & Spaces(0)
                        ActI_str = ActI_str & Spaces(0)
                        Tim5_str = Tim5_str & Spaces(0)
                        Tim6_str = Tim6_str & Spaces(0)
                        Tim7_str = Tim7_str & Spaces(0)
                        Tim8_str = Tim8_str & Spaces(0)
                    End If
                Else
                    Exit Do
                End If
                If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
                dtfromdate = DateAdd("d", 1, dtfromdate)
            Loop 'END OF DATE LOOP
            If Trim(A_Str) = "" And Trim(D_Str) = "" And Trim(L_Str) = "" And Trim(E_str) = "" _
            And Trim(W_Str) = "" And Trim(O_Str) = "" And Trim(p_str) = "" And Trim(S_Str) = "" _
            And Trim(ActO_str) = "" And Trim(ActI_str) = "" And Trim(Tim5_str) = "" And _
            Trim(Tim6_str) = "" And Trim(Tim7_str) = "" And Trim(Tim8_str) = "" Then
            Else
                ConMain.Execute "insert into " & strRepFile & "" & _
                "(Empcode," & strKDate & ",arrstr,depstr,latestr,earlstr,workstr,otstr," & _
                "presabsstr,shfstr,sumlate,sumearly,sumwork,sumextra,ACTRT_O,ACTRT_I," & _
                " TIME5,TIME6,TIME7,TIME8)  values('" & STRECODE & "','" & DTESTR & "','" & _
                A_Str & "','" & D_Str & "','" & L_Str & "','" & E_str & "','" & W_Str & _
                "','" & O_Str & "','" & p_str & "','" & S_Str & "'," & Round(lsum, 2) & _
                "," & esum & "," & wsum & "," & osum & ",'" & ActO_str & "','" & ActI_str & _
                "','" & Tim5_str & "','" & Tim6_str & "','" & Tim7_str & "','" & Tim8_str & "')"
            End If
            lsum = 0: esum = 0: wsum = 0: osum = 0
            dtfromdate = dtFirstDate
        Loop 'END OF EMPLOYEE LOOP
    End If
    pe8Punches = True
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If 'adrsTemp.eof
Exit Function
ERR_P:
    ShowError ("Periodic 8 Punches :: Reports")
    ''Resume Next
End Function
''
''For Mauritius 14-07-2003
Public Function pePermission() As Boolean
On Error GoTo ERR_P
Dim lsum As Single, esum As Single, wsum As Single
Dim A_Str As String, D_Str As String, L_Str As String, E_str As String
Dim W_Str As String, Perm_Str As String, p_str As String, S_Str As String
Dim strGP As String, strTmp As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim DTESTR As String, STRECODE As String
Dim strfile1 As String, strFile2 As String

dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

DTESTR = ""
Do While dtfromdate <= dttodate
    If dtfromdate = dttodate Then
        DTESTR = DTESTR & Day(dtfromdate)
    ElseIf dtfromdate <> dttodate Then
        DTESTR = DTESTR & Day(dtfromdate) & Spaces(Len(Trim(str(Day(dtfromdate)))))
    End If
    dtfromdate = DateAdd("d", 1, dtfromdate)
Loop

If strfile1 = strFile2 Then
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs," & _
    "earlhrs,wrkhrs,presabs," & strfile1 & ".shift," & strfile1 & ".od_from," & strfile1 & _
    ".od_to," & strfile1 & ".aflg," & strfile1 & ".dflg from " & strfile1 & "," & rpTables & _
    " where " & strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & _
    DateCompStr(typRep.strPeriFr) & strDTEnc & " and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(typRep.strPeriTo) & strDTEnc & " " & strSql
Else
    strGP = "select " & strfile1 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs," & _
    "earlhrs,wrkhrs,presabs," & strfile1 & ".shift," & strfile1 & ".od_from," & strfile1 & _
    ".od_to," & strfile1 & ".aflg," & strfile1 & ".dflg from " & strfile1 & "," & rpTables & _
    " where " & strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & _
    DateCompStr(typRep.strPeriFr) & strDTEnc & strSql & _
    " union select " & strFile2 & ".Empcode," & strKDate & ",arrtim,deptim,latehrs," & _
    "earlhrs,wrkhrs,presabs," & strFile2 & ".shift," & strFile2 & ".od_from," & strFile2 & _
    ".od_to," & strFile2 & ".aflg," & strFile2 & ".dflg from " & strFile2 & "," & rpTables & _
    " where empmst.Empcode = " & strFile2 & ".Empcode and " & strKDate & "<=" & strDTEnc & _
    DateCompStr(typRep.strPeriTo) & strDTEnc & " " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select

dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
            STRECODE = adrsTemp!Empcode
            dtfromdate = dtFirstDate
            A_Str = ""
            D_Str = ""
            L_Str = ""
            E_str = ""
            W_Str = ""
            Perm_Str = ""
            p_str = ""
            S_Str = ""
            Do While dtfromdate <= dttodate
                If adrsTemp.EOF Then Exit Do
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtfromdate Then
                        A_Str = A_Str & IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), Spaces(0))
                                                        
                        D_Str = D_Str & IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), Spaces(0))
                                                        
                        L_Str = L_Str & IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                        Spaces(Len(Format(adrsTemp!latehrs, "0.00"))) & Format(adrsTemp!latehrs, "0.00"), Spaces(0))
                                                        
                        E_str = E_str & IIf(Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                        Spaces(Len(Format(adrsTemp!earlhrs, "0.00"))) & Format(adrsTemp!earlhrs, "0.00"), Spaces(0))
                                                        
                        W_Str = W_Str & IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                        Spaces(Len(Format(adrsTemp!wrkHrs, "0.00"))) & Format(adrsTemp!wrkHrs, "0.00"), Spaces(0))
                        
                        ''Permission card Show
                        If adrsTemp!Od_From > 0 And adrsTemp!Od_To > 0 Then
                            strTmp = "O-"
                        End If
                        If adrsTemp!aflg = "1" Then
                            strTmp = strTmp & "L-"
                        ElseIf adrsTemp!aflg = "3" Then
                            strTmp = strTmp & "B-"
                        End If
                        If adrsTemp!Dflg = "2" Then
                            strTmp = strTmp & "E-"
                        End If
                        If Len(Trim(strTmp)) > 0 Then
                            strTmp = Left(strTmp, Len(strTmp) - 1)
                            Perm_Str = Perm_Str & Spaces(Len(strTmp)) & strTmp
                            strTmp = ""
                        Else
                            Perm_Str = Perm_Str & Spaces(0)
                        End If
                        
                        p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                                                        
                        S_Str = S_Str & IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "") & Spaces(Len(adrsTemp!Shift))
                        
                        lsum = TimAdd(IIf(IsNull(lsum), 0, lsum), IIf(IsNull(adrsTemp!latehrs) Or adrsTemp!latehrs <= 0, 0, adrsTemp!latehrs))
                        esum = TimAdd(IIf(IsNull(esum), 0, esum), IIf(IsNull(adrsTemp!earlhrs) Or adrsTemp!earlhrs <= 0, 0, adrsTemp!earlhrs))
                        wsum = TimAdd(IIf(IsNull(wsum), 0, wsum), IIf(IsNull(adrsTemp!wrkHrs), 0, adrsTemp!wrkHrs))
                        
                    ElseIf adrsTemp!Date <> dtfromdate Then
                        A_Str = A_Str & Spaces(0)
                        D_Str = D_Str & Spaces(0)
                        L_Str = L_Str & Spaces(0)
                        E_str = E_str & Spaces(0)
                        W_Str = W_Str & Spaces(0)
                        Perm_Str = Perm_Str & Spaces(0)
                        p_str = p_str & Spaces(0)
                        S_Str = S_Str & Spaces(0)
                    End If
                Else
                    Exit Do
                End If
                If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
                dtfromdate = DateAdd("d", 1, dtfromdate)
            Loop 'END OF DATE LOOP
            If Trim(Perm_Str) = "" Then
            Else
                ConMain.Execute "insert into " & strRepFile & "" & _
                "(Empcode," & strKDate & ",arrstr,depstr,latestr,earlstr,workstr,otstr," & _
                "presabsstr,shfstr,sumlate,sumearly,sumwork)  values('" & STRECODE & _
                "','" & DTESTR & "','" & A_Str & "','" & D_Str & "','" & L_Str & "','" & _
                E_str & "','" & W_Str & "','" & Perm_Str & "','" & p_str & "','" & S_Str & _
                "'," & Round(lsum, 2) & "," & esum & "," & wsum & ")"
            End If
            lsum = 0: esum = 0: wsum = 0
            dtfromdate = dtFirstDate
        Loop 'END OF EMPLOYEE LOOP
    End If
    pePermission = True
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
End If 'adrsTemp.eof
Exit Function
ERR_P:
    ShowError ("Periodic Performance Overtime :: Reports")
    ''Resume Next
End Function

''For mauritius 04-08-31
Public Function peLeaveAvail() As Boolean
On Error GoTo ERR_P
Dim strfile1 As String, strFile2 As String
Dim dtfromdate As Date, dttodate As Date
Dim dtFirstDate As Date, dtLastDate As Date
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

If strfile1 = strFile2 Then
    If FindTable(strfile1) Then
        If adrsEmp.State = 1 Then adrsEmp.Close
        adrsEmp.Open "Select max(" & strKDate & ") as MaxD from " & strfile1, ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsEmp.EOF And adrsEmp.BOF) Then
            dtFirstDate = adrsEmp(0)
        End If
    Else '' If the transaction file is not found
        dtFirstDate = "01/01/2000"
    End If
Else
    If FindTable(strfile1) Then
        If adrsEmp.State = 1 Then adrsEmp.Close
        adrsEmp.Open "Select max(" & strKDate & ") as MaxD from " & strfile1, ConMain, adOpenStatic, adLockOptimistic
        If IsEmpty(adrsEmp(0)) Or IsNull(adrsEmp(0)) Or adrsEmp(0) <= 0 Then
            dtFirstDate = "01/01/2000"
        Else
            dtFirstDate = adrsEmp(0)
        End If
    Else '' If the transaction file is not found
        dtFirstDate = "01/01/2000"
    End If
    If FindTable(strFile2) Then
        If adrsEmp.State = 1 Then adrsEmp.Close
        adrsEmp.Open "Select max(" & strKDate & ") as MaxD from " & strFile2, ConMain, adOpenStatic, adLockOptimistic
        If IsEmpty(adrsEmp(0)) Or IsNull(adrsEmp(0)) Or adrsEmp(0) <= 0 Then
            dtLastDate = "01/01/2000"
        Else
            dtLastDate = adrsEmp(0)
        End If
    Else '' If the transaction file is not found
        dtLastDate = "01/01/2000"
    End If
End If

If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select " & strMon_Trn & ".Empcode,lcode,fromdate,todate,days,trcd  from " & _
strMon_Trn & "," & rpTables & " where trcd =4 and (fromdate between " & strDTEnc & _
DateCompStr(dtfromdate) & strDTEnc & " and " & strDTEnc & DateCompStr(dttodate) & _
strDTEnc & " or todate between " & strDTEnc & DateCompStr(dtfromdate) & strDTEnc & _
" and " & strDTEnc & DateCompStr(dttodate) & strDTEnc & ") and " & strMon_Trn & _
".Empcode = Empmst.Empcode " & strSql & " order by Empmst.Empcode,fromdate", ConMain, adOpenStatic, adLockOptimistic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    Do While Not adrsTemp.EOF
        If adrsTemp("fromdate") <= dtFirstDate Or adrsTemp("fromdate") <= dtLastDate Then
            ''Processed data do not show
        Else
            ConMain.Execute " insert into " & strRepFile & "(Empcode,lcode," & _
            "fromdate,todate,days,trcd ) values('" & adrsTemp("Empcode") & "','" & adrsTemp("lcode") & _
            "'," & strDTEnc & DateCompStr(adrsTemp("Fromdate")) & strDTEnc & "," & strDTEnc & DateCompStr(adrsTemp("ToDate")) & _
            strDTEnc & "," & adrsTemp("days") & ",' ')"
        End If
        adrsTemp.MoveNext
    Loop
End If
peLeaveAvail = True
Exit Function
ERR_P:
    ShowError ("peLeaveAvail :: mdlRep")
    ''Resume Next
End Function

'

Public Function YearCount1(ByVal strFlName As String, ByVal strAbPrS As String, _
ByVal STRECODE As String, ByVal strDateIn As String) As String
On Error GoTo ERR_P
Dim strGY As String

strDateIn = FdtLdt(Month(DateCompDate(strDateIn)), Year(DateCompDate(strDateIn)), "L")
strGY = "select sum(" & strAbPrS & ") from " & strFlName & " where lst_date = " & strDTEnc & _
    DateCompStr(strDateIn) & strDTEnc & " and Empcode=" & "'" & STRECODE & "'"
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGY, ConMain, adOpenStatic
If Not (adrsTemp.EOF Or adrsTemp.BOF) Then
    YearCount1 = IIf(adrsTemp(0) = 0 Or IsNull(adrsTemp(0)) = True, _
        "0", Format(adrsTemp(0), "0.00"))
End If
Exit Function
ERR_P:
    ShowError ("Year Count1 :: Reports")
    YearCount1 = "0"
End Function
Public Function yrManPerf() As Boolean
On Error GoTo ERR_P
yrManPerf = True
Dim strYrMan As String, strLVselect As String
Dim strCurMon As String, strCurYear As String
Dim strFdate As String, strLdate As String
Dim STRECODE As String, bytCnt As Integer
Dim strFileName As String, intCounter As Integer
Dim adrsLvCD As New ADODB.Recordset
intCounter = 10
strFdate = FdtLdt(CByte(pVStar.Yearstart), pVStar.YearSel, "f")
strLdate = FdtLdt(CByte(pVStar.Yearstart) - 1, IIf(pVStar.YearSel = "1", _
typRep.strYear, pVStar.YearSel + 1), "l")

strCurMon = strFdate
strCurYear = Year(DateCompDate(strFdate))

If adrsLvCD.State = 1 Then adrsLvCD.Close
adrsLvCD.Open "select distinct lvcode from leavdesc where lvcode not in ('WO','A ','HL','P ')", ConMain, adOpenStatic
If Not (adrsLvCD.EOF Or adrsLvCD.BOF) = True Then
    For i = 1 To adrsLvCD.RecordCount
      StrLvCD(i) = "'" & adrsLvCD!LvCode & "'"
      strAlv(i) = "{" & "cmd." & adrsLvCD!LvCode & "}"
      adrsLvCD.MoveNext
      Next
End If

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select DISTINCT EMPCODE from " & rpTables & " where " & _
"joindate <= " & strDTEnc & DateCompStr(strLdate) & strDTEnc & strSql & _
" ORDER BY EMPCODE", ConMain, adOpenStatic

strFileName = "lvtrn" & Right(strYearFrom(strFdate), 2)

YearStr = ""
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "select distinct(lvcode) from leavdesc", ConMain, adOpenStatic
If Not (adrsLeave.BOF And adrsLeave.EOF) Then
    strLVselect = ""
    For i = 1 To adrsLeave.RecordCount
        YearStr = YearStr & adrsLeave(0) & Spaces(Len(adrsLeave(0)))
        If i <> adrsLeave.RecordCount Then
            strLVselect = strLVselect & adrsLeave(0) & ","
        Else
            strLVselect = strLVselect & adrsLeave(0)
        End If
        adrsLeave.MoveNext
    Next i
End If
If adrsLeave.State = 1 Then adrsLeave.Close
'*********
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp!Empcode
        strCurMon = strFdate
        strCurYear = Year(DateCompDate(strFdate))
        For i = 1 To 12
            Select Case Month(DateCompDate(strCurMon))
                Case Month(DateCompDate(strCurMon))
                    
                    strFdate = FdtLdt(Month(DateCompDate(strCurMon)), strCurYear, "f")
                    strLdate = FdtLdt(Month(DateCompDate(strCurMon)), strCurYear, "l")
                    
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    adrsTemp.Open "select paiddays,wrk_hrs,night,lt_no,lt_hrs,erl_no,erl_hrs," & strLVselect & _
                    " from " & strFileName & " where Empcode=" & "'" & STRECODE & "'" & " and lst_date=" & _
                    strDTEnc & DateCompStr(strLdate) & strDTEnc, ConMain, adOpenStatic
                    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                        For bytCnt = 7 To adrsTemp.Fields.Count - 1
                            strYrMan = strYrMan & IIf(Not IsNull(adrsTemp(bytCnt)) And adrsTemp(bytCnt) > 0, _
                            Format(adrsTemp(bytCnt), "0.00") & Spaces(Len(Format(adrsTemp(bytCnt), "0.00"))), Spaces(0))
                        Next
                        Dim strGY As String
                        Select Case typOptIdx.bytYer
                            Case 1, 13 'Mandays
                                If Trim(strYrMan) = "" And adrsTemp(0) <= 0 And adrsTemp(1) <= 0 And adrsTemp(2) <= 0 Then
                                Else
                                    strGY = "insert into " & strRepFile & " (Empcode,ystr,yvalstr,pddaysstr,wrkstr,nightstr,[counter]) " & _
                                    " values('" & STRECODE & "'," & "'" & Left(MonthName(Month(DateCompDate(strCurMon))), 3) & _
                                    Right(strCurYear, 2) & "'" & "," & "'" & strYrMan & "'" & "," & _
                                    IIf(Not IsNull(adrsTemp(0)) And adrsTemp(0) > 0, adrsTemp(0), "''") & "," & _
                                    IIf(Not IsNull(adrsTemp(1)) And adrsTemp(1) > 0, adrsTemp(1), "''") & "," & _
                                    IIf(Not IsNull(adrsTemp(2)) And adrsTemp(2) > 0, adrsTemp(2), "''") & "," & intCounter & ")"
                                    intCounter = intCounter + 1
                                End If
                            Case 2 'Performance
                                If Trim(strYrMan) = "" And adrsTemp(0) <= 0 And adrsTemp(1) <= 0 And adrsTemp(2) <= 0 _
                                And adrsTemp(3) <= 0 And adrsTemp(4) <= 0 And adrsTemp(5) <= 0 And adrsTemp(6) <= 0 Then
                                Else
                                    strGY = "insert into " & strRepFile & "(Empcode,ystr,yvalstr,pddaysstr,wrkstr,nightstr,ltno," & _
                                    "latehrs,erno,earlhrs,[counter]) values( '" & STRECODE & "'" & "," & "'" & _
                                    Left(MonthName(Month(DateCompDate(strCurMon))), 3) & Right(strCurYear, 2) & "'," & _
                                    "'" & strYrMan & "','" & IIf(Not IsNull(adrsTemp(0)) And adrsTemp(0) > 0, adrsTemp(0), " ") & "','" & _
                                    IIf(Not IsNull(adrsTemp(1)) And adrsTemp(1) > 0, adrsTemp(1), "") & "','" & _
                                    IIf(Not IsNull(adrsTemp(2)) And adrsTemp(2) > 0, adrsTemp(2), "") & "','" & _
                                      IIf(Not IsNull(adrsTemp(3)) And adrsTemp(3) > 0, adrsTemp(3), "") & "','" & _
                                    IIf(Not IsNull(adrsTemp(4)) And adrsTemp(4) > 0, adrsTemp(4), "") & "','" & _
                                    IIf(Not IsNull(adrsTemp(5)) And adrsTemp(5) > 0, adrsTemp(5), "") & "','" & _
                                    IIf(Not IsNull(adrsTemp(6)) And adrsTemp(6) > 0, adrsTemp(6), "") & "'," & intCounter & ")"
                                    intCounter = intCounter + 1
                                End If
                        End Select
                        If Trim(strGY) <> "" Then ConMain.Execute strGY
                        strYrMan = ""
                    End If
            End Select
            strCurMon = CStr(DateAdd("m", 1, DateCompDate(strCurMon)))
            If Month(DateCompDate(strCurMon)) < Month(DateCompDate(strLdate)) Then
                strCurYear = Year(DateCompDate(strFdate)) + 1
            End If
        Next i
        strCurYear = Val(typRep.strYear)
        strFdate = FdtLdt(Month(DateCompDate(strCurMon)), strCurYear, "f")
        adrsEmp.MoveNext
        If bytBackEnd = 2 Then Sleep (100)
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Yearly Man Performance :: Reports")
'    Resume Next
    yrManPerf = False
End Function

Public Function yrLeaveBAl() As Boolean
On Error GoTo Err
Dim strLeaveCode As String
Dim strY As String
Dim adrsT As New ADODB.Recordset
Dim strI() As Variant
Dim j As Integer
Dim strIForT As String
Dim strheader As String
strY = "lvbal" & Right(typRep.strYear, 2)
strLeaveCode = GetLEaveCodeFromBal(strY, ",")
strLeaveCode = Left(strLeaveCode, Len(strLeaveCode) - 1)

Set adrsT = OpenRecordSet("SELECT " & strLeaveCode & " FROM " & strY & "," & rpTables & _
" WHERE " & strY & ".Empcode=empmst.Empcode " & strSql & "")

If FindTable("TmpLeavBal") Then
    ConMain.Execute "DELETE FROM TmpLeavBal"
Else
    ConMain.Execute "CREATE TABLE TmpLeavBal(strT VARCHAR(200))"
End If
'for header
If FindTable("TmpLeavBalH") Then
    ConMain.Execute "DELETE FROM TmpLeavBalH"
Else
    ConMain.Execute "CREATE TABLE TmpLeavBalH(strT VARCHAR(200))"
End If
strI = adrsT.GetRows(adrsT.RecordCount)
strheader = Replace(strLeaveCode, strY & ".", "")
strheader = Replace(strheader, ",", Space(4))
'for inserting
ConMain.Execute "INSERT INTO TmpLeavBalH(strT) " & _
" VALUES ('" & strheader & "')"
For i = 0 To UBound(strI, 2)
    strIForT = ""
    For j = 0 To UBound(strI, 1)
        strIForT = strIForT & IIf(IsNull(strI(j, i)), Space(2), strI(j, i)) & _
        Space(4) + IIf(j = 0, Space((Len("Empcode") + Len(Space(4))) - Len(pVStar.CodeSize)), Space(0))
'        strIForT = strIForT & IIf(IsNull(strI(j, i)), Space(4), strI(j, i)) & _
'        Space(4) & Space(IIf(Len(strI(j, i)) < 4, 4 - Len(strI(j, i)), 0))
    Next
    Debug.Print strIForT
    ConMain.Execute "INSERT INTO TmpLeavBal(strT) " & _
    " VALUES ('" & strIForT & "')"
Next
yrLeaveBAl = True
Exit Function
Err:
    yrLeaveBAl = False
    Call ShowError("Error in yrLeaveBAl")
End Function

Public Function yrLeaveInfo() As Boolean
On Error GoTo ERR_P
yrLeaveInfo = True
Dim strYrFrm As String
Dim STRECODE As String, strFileName As String
Dim strFromDate As String, strToDate As String

strFromDate = FdtLdt(CByte(pVStar.Yearstart), typRep.strYear, "f")
strToDate = FdtLdt(CByte(IIf(pVStar.Yearstart = 1, "12", pVStar.Yearstart - 1)), IIf(pVStar.Yearstart = "1", _
    typRep.strYear, CStr(Val(typRep.strYear) + 1)), "l")

strFileName = "lvinfo" & Right(strYearFrom(CStr(strFromDate)), 2)
  


If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select DISTINCT EMPCODE from " & rpTables & " where " & _
"joindate<=" & strDTEnc & DateCompStr(strToDate) & strDTEnc & " " & strSql & " order by Empcode" _
, ConMain, adOpenStatic

If Not (adrsEmp.BOF And adrsEmp.EOF) Then
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp!Empcode
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select trcd,fromdate,todate,days,lcode from " & strFileName & _
        " where Empcode='" & STRECODE & "' order by lcode,fromdate", ConMain, adOpenStatic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not adrsTemp.EOF
                With ConMain
                    Select Case adrsTemp!trcd
                        Case 1:
                            strYrFrm = NewCaptionTxt("M7022", adrsMod)
                            .Execute "insert into " & strRepFile & "(Empcode,fromlv,creditlv,fromdate," & _
                            "todate,lcode,trcd) values (" & "'" & STRECODE & "'," & "'" & strYrFrm & "'" & _
                            "," & adrsTemp("days") & "," & strDTEnc & DateCompStr(adrsTemp!FromDate) & _
                            strDTEnc & "," & strDTEnc & DateCompStr(adrsTemp!ToDate) & strDTEnc & "," & _
                            "'" & adrsTemp!LCode & "'" & "," & adrsTemp!trcd & ")"
                        Case 2:
                            strYrFrm = NewCaptionTxt("M7023", adrsMod)
                            .Execute "insert into " & strRepFile & "(Empcode,fromlv,creditlv,fromdate," & _
                            "todate,lcode,trcd) values (" & "'" & STRECODE & "'" & ",'" & strYrFrm & "'" & _
                            "," & adrsTemp("days") & "," & strDTEnc & DateCompStr(adrsTemp!FromDate) & _
                            strDTEnc & "," & strDTEnc & DateCompStr(adrsTemp!ToDate) & strDTEnc & "," & _
                            "'" & adrsTemp!LCode & "'" & "," & adrsTemp!trcd & ")"
                        Case 3:
                            strYrFrm = NewCaptionTxt("M7024", adrsMod)
                            .Execute "insert into " & strRepFile & "(Empcode,fromlv,availlv,fromdate," & _
                            "todate,lcode,trcd) values (" & "'" & STRECODE & "'," & "'" & strYrFrm & "'" & _
                            ",'" & adrsTemp("days") & "'," & strDTEnc & DateCompStr(adrsTemp!FromDate) & _
                            strDTEnc & "," & strDTEnc & DateCompStr(adrsTemp!ToDate) & strDTEnc & "," & _
                            "'" & adrsTemp!LCode & "'" & ",'" & adrsTemp!trcd & "')"
                        Case 4:
                            strYrFrm = DateDisp(adrsTemp!FromDate)
                            .Execute "insert into " & strRepFile & "(Empcode,fromlv,availlv,fromdate," & _
                            "todate,lcode,trcd) values (" & "'" & STRECODE & "'" & ",'" & strYrFrm & "'" & _
                            "," & adrsTemp("days") & "," & strDTEnc & DateCompStr(adrsTemp!FromDate) & _
                            strDTEnc & "," & strDTEnc & DateCompStr(adrsTemp!ToDate) & strDTEnc & "," & _
                            "'" & adrsTemp!LCode & "'" & "," & adrsTemp!trcd & ")"
                        Case 6:
                            strYrFrm = NewCaptionTxt("M7025", adrsMod)
                            .Execute "insert into " & strRepFile & "(Empcode,fromlv,availlv,fromdate," & _
                            "todate,lcode,trcd) values (" & "'" & STRECODE & "'," & "'" & strYrFrm & "'" & _
                            ",'" & adrsTemp("days") & "'," & strDTEnc & DateCompStr(adrsTemp!FromDate) & _
                            strDTEnc & "," & strDTEnc & DateCompStr(adrsTemp!ToDate) & strDTEnc & "," & _
                            "'" & adrsTemp!LCode & "'" & ",'" & adrsTemp!trcd & "')"
                        Case 7:
                            strYrFrm = NewCaptionTxt("M7025", adrsMod)
                            .Execute "insert into " & strRepFile & "(Empcode,fromlv,availlv,fromdate," & _
                            "todate,lcode,trcd) values (" & "'" & STRECODE & "'," & "'" & strYrFrm & "'" & _
                            ",'" & adrsTemp("days") & "'," & strDTEnc & DateCompStr(adrsTemp!FromDate) & _
                            strDTEnc & "," & strDTEnc & DateCompStr(adrsTemp!ToDate) & strDTEnc & "," & _
                            "'" & adrsTemp!LCode & "'" & ",'" & adrsTemp!trcd & "')"
                            
                    End Select
                End With
                strYrFrm = ""
                adrsTemp.MoveNext
            Loop
            If adrsTemp.State = 1 Then adrsTemp.Close
        End If

        adrsEmp.MoveNext
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Yearly Leave Info :: Reports")
    yrLeaveInfo = False
End Function

Public Sub SetMSF1Cap(Optional bytNum As Byte = 0)
On Error GoTo ERR_P
Dim strA As String, strB As String, strC As String, strD As String, strE As String
Dim strF As String
If CatFlag = True Then
With frmReports
    .MSF1.Col = 0
    .MSF1.Row = 0
    .MSF1.Redraw = True
    .MSF1.ForeColor = vbBlue
    Select Case bytAction
        Case 1: strF = "Mail"
        Case 2: strF = "View"
        Case 3: strF = "Print"
        Case 4: strF = "Print to File"
    End Select
    strA = "   Checking Validations..."
    strB = "   Processing Valid Data..."
    strC = "   Executing Query..."
    strD = "   Operation Aborted"
    strE = "   Preparing Report to " & strF
    Select Case bytNum
        Case 0:    .MSF1.Text = "   " & Err.Description & strD: .MSF1.ForeColor = vbRed
        Case 1:    .MSF1.Text = "   Daily Reports"
        Case 2:    .MSF1.Text = "   Weekly Reports"
        Case 3:    .MSF1.Text = "   Monthly Reports"
        Case 4:    .MSF1.Text = "   Yearly Reports"
        Case 5:    .MSF1.Text = "   Masters Reports"
        Case 6:    .MSF1.Text = "   Periodic Reports"
        Case 7:    .MSF1.Text = strA
        Case 8:    .MSF1.Text = strB
        Case 9:    .MSF1.Text = strC
        Case 10:   .MSF1.Text = strD: .MSF1.ForeColor = vbRed
        Case 11:   .MSF1.Text = strE
        Case 12:   .MSF1.Text = "    Sending Mail to "
        Case 13:   .MSF1.Text = "   "
        Case 14:   .MSF1.Text = "   "
        Case 15:   .MSF1.Text = "   "
        Case 16:   .MSF1.Text = "   "
        Case 17:   .MSF1.Text = "   "
        Case 18:   .MSF1.Text = "   "
        Case 19:   .MSF1.Text = "   "
        Case 20:   .MSF1.Text = "   "
    End Select
    .MSF1.Refresh
End With
Else
End If
Exit Sub
ERR_P:
    ShowError ("Set Message Caption :: Reports")
    'Resume Next
End Sub

'  to modify in crystal
Public Function monPerfCry()
On Error GoTo ERR_P
monPerfCry = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim ArrArray(31), DepArray(31), LateArray(31), EarlArray(31), WrkArray(31), OTarray(31) As Single
Dim PresArray(31), ShfArray(31) As String
Dim DateArray(31) As Integer
Dim ArrFld As Variant, ArrFld1 As Variant
Dim arrstr As Variant, ArrStr1 As Variant
Dim tmpStr As String

Dim STRECODE As String, strTrnFile As String, strDateS As String
Dim valFld As String, ValFld1 As String
Dim i, j
Dim strFld As String, strFld1 As String

dtfromdate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "f"))
dttodate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))

strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")
If typOptIdx.bytMon = 0 Then
    tmpStr = ""
ElseIf typOptIdx.bytMon = 5 Then
    tmpStr = " and ovtim > 0 "
ElseIf typOptIdx.bytMon = 10 Then
    tmpStr = " and latehrs > 0 "
ElseIf typOptIdx.bytMon = 11 Then
    tmpStr = " and earlhrs > 0 "
End If

If adrsTemp.State = 1 Then adrsTemp.Close
   adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
   "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
   " empmst.Empcode = " & strTrnFile & ".Empcode " & tmpStr & strSql & _
   " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic

strFld = "Empcode"

If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
    'strDateS = monDateStr(Day(dtToDate)) 'Assigning date string
    Do While Not (adrsTemp.EOF) 'And dtTempDate < dtToDate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    DateArray(Day(adrsTemp!Date)) = Day(dtTempDate)
                    ArrArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                                Format(adrsTemp!arrtim, "0.00"), 0)
                                     
                    DepArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                                Format(adrsTemp!deptim, "0.00"), 0)
                    
                    ShfArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "")
                    
                    LateArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!latehrs) And adrsTemp!latehrs > 0, _
                              Format(adrsTemp!latehrs, "0.00"), 0)
                    
                    EarlArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!earlhrs) And adrsTemp!earlhrs > 0, _
                              Format(adrsTemp!earlhrs, "0.00"), 0)
                    
                    WrkArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                             Format(adrsTemp!wrkHrs, "0.00"), 0)
                        
                      
                        If adrsTemp("OTConf") = "Y" Then
                            OTarray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!ovtim) And adrsTemp!ovtim > 0, _
                             Format(adrsTemp!ovtim, "0.00"), 0)
                        Else
                            OTarray(Day(adrsTemp!Date)) = 0
                        End If
                        
                        If (Left(adrsTemp!presabs, 2) = pVStar.WosCode Or _
                            Left(adrsTemp!presabs, 2) = pVStar.HlsCode) And adrsTemp!arrtim > 0 Then
                                PresArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs & "p", "")
                        Else
                                PresArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")

                        End If
                    
                ElseIf adrsTemp!Date <> dtTempDate Then
                     DateArray(Day(adrsTemp!Date)) = 0
                    ArrArray(Day(dtTempDate)) = 0
                    DepArray(Day(dtTempDate)) = 0
                    ShfArray(Day(dtTempDate)) = ""
                    LateArray(Day(dtTempDate)) = 0
                    EarlArray(Day(dtTempDate)) = 0
                    WrkArray(Day(dtTempDate)) = 0
                    PresArray(Day(dtTempDate)) = ""
                    OTarray(Day(dtTempDate)) = 0
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        valFld = "'" & STRECODE & "'"
        ValFld1 = "'" & STRECODE & "'"
        
        
        ' Performance
        If typOptIdx.bytMon <> 10 And typOptIdx.bytMon <> 11 Then
        ' Splitting the Data for MS access Database and getting the data in two tables
        For j = 0 To 6
        ArrFld = Array(ArrArray(), DepArray(), LateArray(), EarlArray(), WrkArray(), OTarray(), DateArray(), PresArray(), ShfArray())

            For i = 1 To Day(dttodate)
                If (ArrFld(6)(i)) <> 0 Then
                    valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, ArrFld(j)(i))
                End If
            Next i
            
        Next j
        
        For j = 7 To 8
        ArrFld1 = Array(ArrArray(), DepArray(), LateArray(), EarlArray(), WrkArray(), OTarray(), DateArray(), PresArray(), ShfArray())

            For i = 1 To Day(dttodate)
                If (ArrFld1(6)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)), "", ArrFld1(j)(i)) & "'"
                End If
            Next i
            
        Next j
        End If
        
        ' Late Arrival
        If typOptIdx.bytMon = 10 Then
        For j = 0 To 3
        ArrFld = Array(ArrArray(), DepArray(), LateArray(), DateArray(), ShfArray())

            For i = 1 To Day(dttodate)
                If (ArrFld(3)(i)) <> 0 Then
                    valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, ArrFld(j)(i))
                End If
            Next i
            
        Next j
        For j = 4 To 4
        ArrFld1 = Array(ArrArray(), DepArray(), LateArray(), DateArray(), ShfArray())

            For i = 1 To Day(dttodate)
                If (ArrFld1(3)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)), "", ArrFld1(j)(i)) & "'"
                End If
            Next i
            
        Next j
        End If
        
        ' Early Arrival
        If typOptIdx.bytMon = 11 Then
        For j = 0 To 3
        ArrFld = Array(ArrArray(), DepArray(), EarlArray(), DateArray(), ShfArray())

            For i = 1 To Day(dttodate)
                If (ArrFld(3)(i)) <> 0 Then
                    valFld = valFld & "," & IIf(IsEmpty(ArrFld(j)(i)), 0, ArrFld(j)(i))
                End If
            Next i
            
        Next j
        For j = 4 To 4
        ArrFld1 = Array(ArrArray(), DepArray(), EarlArray(), DateArray(), ShfArray())

            For i = 1 To Day(dttodate)
                If (ArrFld1(3)(i)) <> 0 Then
                    ValFld1 = ValFld1 & ",'" & IIf(IsEmpty(ArrFld1(j)(i)), "", ArrFld1(j)(i)) & "'"
                End If
            Next i
            
        Next j
        End If
        
        If typOptIdx.bytMon = 10 Then
            arrstr = Array("Arr", "Dep", "Late", "dt")
            ArrStr1 = Array("shf")
        ElseIf typOptIdx.bytMon = 11 Then
            arrstr = Array("Arr", "Dep", "Earl", "dt")
            ArrStr1 = Array("shf")
        Else
            arrstr = Array("Arr", "Dep", "Late", "Earl", "Work", "OT", "Dt")
            ArrStr1 = Array("Rem", "shf")
        End If
        strFld = "Empcode"
        strFld1 = strFld
        
        For j = 0 To UBound(arrstr)
            For i = 1 To Day(dttodate)
                If typOptIdx.bytMon = 10 Or typOptIdx.bytMon = 11 Then
                    If (ArrFld(3)(i)) <> 0 Then
                        strFld = strFld & "," & arrstr(j) & (i)
                    End If
                Else
                    If (ArrFld(6)(i)) <> 0 Then
                        strFld = strFld & "," & arrstr(j) & (i)
                    End If
    
                End If
            Next i
        Next j
        
        '************************ : this is for the new table
        
        For j = 0 To UBound(ArrStr1)
            For i = 1 To Day(dttodate)
                If typOptIdx.bytMon = 10 Or typOptIdx.bytMon = 11 Then
                    If (ArrFld(3)(i)) <> 0 Then
                        strFld1 = strFld1 & "," & ArrStr1(j) & (i)
                    End If
                Else
                    If (ArrFld1(6)(i)) <> 0 Then
                        strFld1 = strFld1 & "," & ArrStr1(j) & (i)
                    End If
                    
                    
                    
    
                End If
            Next i
        Next j
        ': Data is been splitted and entered into two table
        
        ConMain.Execute "insert into " & strRepFile & "(" & strFld & ") values(" & valFld & ")"
         ConMain.Execute "insert into " & strRepMfile & "(" & strFld1 & ") values(" & ValFld1 & ")"
       
          Erase ArrArray: Erase DepArray: Erase LateArray: Erase OTarray: Erase DateArray: Erase WrkArray: Erase EarlArray
         
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Monthly Performance :: Reports")
    Resume Next
    monPerfCry = False
End Function


Public Function monFormTwelve()
On Error GoTo ERR_P
    monFormTwelve = True
    Dim adrsLvCD As New ADODB.Recordset
    Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
    Dim ShfArray(31) As String
    Dim DateArray(31) As Integer
    Dim ArrFld As Variant
    Dim arrstr As Variant
    Dim strqry As String
    Dim STRECODE As String, strTrnFile As String, strDateS As String
    Dim valFld As String
    Dim i, j
    Dim strFld As String
    Dim strShiftAll As String
    Dim strLastDate As String
    dtfromdate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), _
        typRep.strMonYear, "f"))
    dttodate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), _
        typRep.strMonYear, "l"))
    strlstdt = Day(dttodate)

    strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")

    strqry = "SELECT empmst.styp,empmst.Empcode," & strKDate & _
        ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf FROM " & strTrnFile & "," & _
        rpTables & " WHERE " & " empmst.Empcode = " & strTrnFile & ".Empcode " & _
        strSql & " ORDER BY empmst.Empcode," & strKDate & ""
    
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open strqry, ConMain, adOpenStatic

    strFld = "Empcode"

    arrstr = Array("Rem")

    For j = 0 To UBound(arrstr)
        For i = 1 To Day(dttodate)
            strFld = strFld & "," & arrstr(j) & (i)
        Next i
    Next j
    'Used to Shift Allowances
    strFld = strFld & ",Shf1"
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
        dtTempDate = dtfromdate
        Do While Not (adrsTemp.EOF) 'And dtTempDate < dtToDate
            STRECODE = adrsTemp!Empcode
            dtTempDate = dtfromdate
            
            Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
                If adrsTemp!Empcode = STRECODE Then
                    If adrsTemp!Date = dtTempDate Then
                        DateArray(Day(adrsTemp!Date)) = Day(dtTempDate)
                        If GetFlagStatus("FORM25") Then
                        ShfArray(Day(adrsTemp!Date)) = _
                                    IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")
                        Else
                        ShfArray(Day(adrsTemp!Date)) = _
                                    IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "")
                        End If
                        If adrsTemp.Fields("styp") = "R" Then
                            strShiftAll = strlstdt
                        Else
                            strShiftAll = "0.00"
                        End If
                    ElseIf adrsTemp!Date <> dtTempDate Then
                         DateArray(Day(adrsTemp!Date)) = Day(dtTempDate)
                         ShfArray(Day(adrsTemp!Date)) = ""
                    End If
                Else
                    Exit Do
                End If
                If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
                dtTempDate = DateAdd("d", 1, dtTempDate)
            Loop
            valFld = "'" & STRECODE & "'"
            j = 0
            ArrFld = Array(ShfArray())
            For i = 1 To Day(dttodate)
                valFld = valFld & ",'" & IIf(IsEmpty(ArrFld(j)(i)), "", ArrFld(j)(i)) & "'"
            Next i
            
            ConMain.Execute "INSERT INTO " & _
                strRepMfile & "(" & strFld & ") VALUES(" & valFld & _
                ",'" & strShiftAll & "')"
            Erase ShfArray
        Loop
    End If

Exit Function
ERR_P:
    ShowError ("Monthly monFormTwelve :: Reports")
    monFormTwelve = False
End Function

'  to modify in crystal
Public Function monMuPACry()
On Error GoTo ERR_P
monMuPACry = True
Dim adrsLvCD As New ADODB.Recordset
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim ArrArray(31), DepArray(31), LateArray(31), EarlArray(31), WrkArray(31), OTarray(31) As Single
Dim PresArray(31), ShfArray(31) As String
Dim DateArray(31) As Integer
Dim ArrFld As Variant
Dim arrstr As Variant
Dim strqry As String
Dim STRECODE As String, strTrnFile As String, strDateS As String
Dim valFld As String
Dim i, j
Dim strFld As String
Dim SecShft As Single

If adrsLvCD.State = 1 Then adrsLvCD.Close
adrsLvCD.Open "select distinct lvcode from leavdesc where lvcode not in ('WO','A ','HL','P ')", ConMain, adOpenStatic
If Not (adrsLvCD.EOF Or adrsLvCD.BOF) = True Then
    For i = 1 To adrsLvCD.RecordCount
      StrLvCD(i) = "'" & adrsLvCD!LvCode & "'"
      strAlv(i) = "{" & "cmd." & adrsLvCD!LvCode & "}"
      adrsLvCD.MoveNext
      Next
 End If
 'Report.FormulaFields.GetItemByName(j).Value = strAlv(i)
dtfromdate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "f"))
dttodate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))
strlstdt = Day(dttodate)

strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")

If typOptIdx.bytMon = 1 Or typOptIdx.bytMon = 2 Or typOptIdx.bytMon = 19 Or typOptIdx.bytMon = 23 Or typOptIdx.bytMon = 25 Or typOptIdx.bytMon = 36 Then
        strqry = "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
                "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
                " empmst.Empcode = " & strTrnFile & ".Empcode " & strSql & _
                " order by empmst.Empcode," & strKDate & ""

ElseIf typOptIdx.bytMon = 3 Then
strqry = "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
                "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
                " empmst.Empcode = " & strTrnFile & ".Empcode and (arrtim > 0 And (presabs=" & "'" & ReplicateVal(pVStar.PrsCode, 2) & "'" & _
        " or " & LeftStr("presabs") & "=" & "'" & pVStar.PrsCode & "'" & " or " & RightStr("presabs") & "=" & _
        "'" & pVStar.PrsCode & "'" & "))" & strSql & _
                " order by empmst.Empcode," & strKDate & ""
                
ElseIf typOptIdx.bytMon = 4 Then
strqry = "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
                "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
                " empmst.Empcode = " & strTrnFile & ".Empcode and (presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
        " or " & LeftStr("presabs") & "=" & "'" & pVStar.AbsCode & "'" & " or " & RightStr("presabs") & "=" & _
        "'" & pVStar.AbsCode & "'" & ")" & strSql & _
                " order by empmst.Empcode," & strKDate & ""
                                
End If

If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strqry, ConMain, adOpenStatic

strFld = "Empcode"

arrstr = Array("Rem")

For j = 0 To UBound(arrstr)
    For i = 1 To Day(dttodate)
        strFld = strFld & "," & arrstr(j) & (i)
    Next i
Next j

If typOptIdx.bytMon = 19 Then
arrstr = Array("Shf")
For j = 0 To UBound(arrstr)
    For i = 1 To Day(dttodate)
        strFld = strFld & "," & arrstr(j) & (i)
    Next i
Next j
End If
'end by
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
    Do While Not (adrsTemp.EOF) 'And dtTempDate < dtToDate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    DateArray(Day(adrsTemp!Date)) = Day(dtTempDate)

                        If typOptIdx.bytMon = 19 Then
                            ShfArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, "")
                        End If
                                 
                        If (Left(adrsTemp!presabs, 2) = pVStar.WosCode Or _
                            Left(adrsTemp!presabs, 2) = pVStar.HlsCode) And adrsTemp!arrtim > 0 Then
                                If Not (adrsTemp!presabs) = pVStar.HlsCode + pVStar.PrsCode Then ' 2-07-09
                                    PresArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs & "p", "")
                                Else
                                    PresArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")
                                End If
                        Else
                            PresArray(Day(adrsTemp!Date)) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")
                        End If

                ElseIf adrsTemp!Date <> dtTempDate Then
                     DateArray(Day(adrsTemp!Date)) = Day(dtTempDate)
                     PresArray(Day(dtTempDate)) = ""
                     ShfArray(Day(adrsTemp!Date)) = ""
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
         valFld = "'" & STRECODE & "'"
        
        j = 0
        ArrFld = Array(PresArray())
            For i = 1 To Day(dttodate)
                valFld = valFld & ",'" & IIf(IsEmpty(ArrFld(j)(i)), "", ArrFld(j)(i)) & "'"
            Next i
        
        If typOptIdx.bytMon = 19 Then
            ArrFld = Array(ShfArray())
            For i = 1 To Day(dttodate)
                 valFld = valFld & ",'" & IIf(IsEmpty(ArrFld(j)(i)), "", ArrFld(j)(i)) & "'"
            Next i
        End If
        'end by
        
         ConMain.Execute "insert into " & strRepMfile & "(" & strFld & ") values(" & valFld & ")"
        Erase ArrArray: Erase DepArray: Erase LateArray: Erase OTarray: Erase DateArray: Erase WrkArray: Erase EarlArray: Erase PresArray
        Erase ShfArray: SecShft = 0
    Loop
    
End If

Exit Function
ERR_P:
    ShowError ("Monthly Attendance :: Reports")
    monMuPACry = False
End Function
    
' For Jublient Client    21-03
Public Function monJublientReport()
On Error GoTo ERR_P
monJublientReport = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim strTrnFile As String, strqry As String, STRECODE As String, strFld As String
Dim valFld As String, strAbsentDays As String, strLateDays As String, strEarlDays As String
Dim absCnt As Double
Dim bytLatcnt As Byte, bytErlcnt As Byte
Dim sngLathrs As Single, sngErlhrs As Single
Dim strTrnFile2 As String, RepIndex As Integer

If bytRepMode = 3 Then
    strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")
    RepIndex = typOptIdx.bytMon
ElseIf bytRepMode = 6 Then
    strTrnFile = MakeName(MonthName(Month(DateCompDate(typRep.strPeriFr))), Year(DateCompDate(typRep.strPeriFr)), "trn")
    strTrnFile2 = MakeName(MonthName(Month(DateCompDate(typRep.strPeriTo))), Year(DateCompDate(typRep.strPeriFr)), "trn")
    RepIndex = typOptIdx.bytPer
End If

Select Case RepIndex
    Case 26, 19 '19 Add By    19-05
        strFld = "Empcode,MndateStr,PresAbsStr"
        strqry = "select empmst.Empcode," & strKDate & ",presabs from " & strTrnFile & "," & rpTables & " where " & _
                " empmst.Empcode = " & strTrnFile & ".Empcode and (presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
                " or " & LeftStr("presabs") & "=" & "'" & pVStar.AbsCode & "'" & " or " & RightStr("presabs") & "=" & _
                "'" & pVStar.AbsCode & "'" & ")" & strSql & ""
       If RepIndex = 19 Then
            strqry = "select empmst.Empcode," & strKDate & ",presabs from " & strTrnFile & "," & rpTables & " where " & _
                " empmst.Empcode = " & strTrnFile & ".Empcode and (presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
                " or " & LeftStr("presabs") & "=" & "'" & pVStar.AbsCode & "'" & " or " & RightStr("presabs") & "=" & _
                "'" & pVStar.AbsCode & "'" & ")" & strSql & " and " & strTrnFile & "." & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & strDTEnc & ""
            If strTrnFile = strTrnFile2 Then
                strqry = strqry + "  and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " "
            Else
                strqry = strqry + "  Union select empmst.Empcode," & strKDate & ",presabs from " & strTrnFile2 & "," & rpTables & " where " & _
                     " empmst.Empcode = " & strTrnFile2 & ".Empcode and (presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
                     " or " & LeftStr("presabs") & "=" & "'" & pVStar.AbsCode & "'" & " or " & RightStr("presabs") & "=" & _
                     "'" & pVStar.AbsCode & "'" & ")" & strSql & "  and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " "
            End If
        End If
        Select Case bytBackEnd
            Case 1, 2 ''SQLServer,MS-Access
                strqry = strqry + " order by empmst.Empcode," & strKDate & ""
            Case 3    '' ORACLE
                strqry = strqry + " order by Empcode," & strKDate & ""
        End Select

        If adrsTemp.State = 1 Then adrsTemp.Close
            adrsTemp.Open strqry, ConMain, adOpenStatic
        
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            adrsTemp.MoveFirst
            Do While Not (adrsTemp.EOF)
                STRECODE = adrsTemp!Empcode: absCnt = 0: strAbsentDays = ","
                Do While Not (adrsTemp.EOF)
                    If (adrsTemp!Empcode = STRECODE) Then
                        If adrsTemp!presabs = ReplicateVal(pVStar.AbsCode, 2) Then
                            absCnt = absCnt + 1
                            strAbsentDays = strAbsentDays & Day(adrsTemp!Date) & ","
                        ElseIf (Left(adrsTemp!presabs, 2) = pVStar.AbsCode) Or (Right(adrsTemp!presabs, 2) = pVStar.AbsCode) Then
                            absCnt = absCnt + 0.5
                            strAbsentDays = strAbsentDays & Day(adrsTemp!Date) & ","
                        End If
                        adrsTemp.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                strAbsentDays = Mid(strAbsentDays, 2, Len(strAbsentDays) - 2)
                valFld = "'" & STRECODE & "','" & absCnt & "','" & strAbsentDays & "'"
                ConMain.Execute "insert into " & strRepFile & "(" & strFld & ") values(" & valFld & ")"
            Loop
        Else
            Call SetMSF1Cap(10)
            monJublientReport = False
            MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
        End If
    
    Case 27, 20 'late/ea rly days report     23-03 '20 index added by  19-05
        bytLatcnt = 0: bytErlcnt = 0: sngLathrs = 0: sngErlhrs = 0
        strqry = "select distinct empmst.Empcode,latehrs,earlhrs," & _
            "" & strKDate & " from " & strTrnFile & "," & rpTables & " where " & strTrnFile & ".Empcode = empmst.Empcode and ((" & _
            strTrnFile & ".latehrs > 0) OR (" & strTrnFile & ".earlhrs > 0)) " & strSql & ""
        
        If RepIndex = 20 Then
            strqry = "select distinct empmst.Empcode,latehrs,earlhrs," & _
                "" & strKDate & " from " & strTrnFile & "," & rpTables & " where " & strTrnFile & ".Empcode = empmst.Empcode and ((" & _
                strTrnFile & ".latehrs > 0) OR (" & strTrnFile & ".earlhrs > 0)) " & strSql & " and " & strTrnFile & "." & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & strDTEnc & ""
            If strTrnFile = strTrnFile2 Then
                strqry = strqry + "  and " & strTrnFile & "." & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " "
            Else
                strqry = strqry + "  Union select distinct empmst.Empcode,latehrs,earlhrs," & _
                    "" & strKDate & " from " & strTrnFile2 & "," & rpTables & " where " & strTrnFile2 & ".Empcode = empmst.Empcode and ((" & _
                    strTrnFile2 & ".latehrs > 0) OR (" & strTrnFile2 & ".earlhrs > 0)) " & strSql & "  and " & strTrnFile2 & "." & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " "
            End If
        End If
        Select Case bytBackEnd
            Case 1, 2 ''SQLServer,MS-Access
                strqry = strqry + " order by empmst.Empcode," & strKDate & ""
            Case 3    '' ORACLE
                strqry = strqry + " order by Empcode," & strKDate & ""
        End Select
        
        If adrsTemp.State = 1 Then adrsTemp.Close
            adrsTemp.Open strqry, ConMain, adOpenStatic
        
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not adrsTemp.EOF
                bytLatcnt = 0: bytErlcnt = 0: sngLathrs = 0: sngErlhrs = 0: strLateDays = ",": strEarlDays = ","
                STRECODE = adrsTemp!Empcode
                Do While STRECODE = adrsTemp!Empcode And Not adrsTemp.EOF
                    If adrsTemp!latehrs > 0 Then
                        sngLathrs = TimAdd(sngLathrs, adrsTemp!latehrs)
                        bytLatcnt = bytLatcnt + 1
                        strLateDays = strLateDays & Day(adrsTemp!Date) & ","
                    End If
                    If adrsTemp!earlhrs > 0 Then
                        sngErlhrs = TimAdd(sngErlhrs, adrsTemp!earlhrs)
                        bytErlcnt = bytErlcnt + 1
                        strEarlDays = strEarlDays & Day(adrsTemp!Date) & ","
                    End If
                    adrsTemp.MoveNext
                    If adrsTemp.EOF Then Exit Do
                Loop
                If strLateDays = "," Then
                    strLateDays = ""
                Else
                    strLateDays = Mid(strLateDays, 2, Len(strLateDays) - 2)
                End If
                If strEarlDays = "," Then
                    strEarlDays = ""
                Else
                    strEarlDays = Mid(Mid(strEarlDays, 1, Len(strEarlDays) - 1), 2)
                End If
                ConMain.Execute " insert into " & strRepFile & _
                    "(Empcode,MndateStr,PresAbsStr,LeaveStr,PDaysStr,OtStr,WrkStr) values('" & _
                    STRECODE & "','" & bytLatcnt & "','" & IIf(sngLathrs > 0, Format(sngLathrs, "0.00"), Empty) & _
                    "','" & strLateDays & "','" & bytErlcnt & "','" & IIf(sngErlhrs > 0, Format(sngErlhrs, "0.00"), Empty) & _
                    "','" & strEarlDays & "')"
             Loop
        Else
            Call SetMSF1Cap(10)
            monJublientReport = False
            MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
        End If
        
        Case 28:    'Absenteeism   28-03
            Dim rsEmp As New ADODB.Recordset
            Dim absent As Double
            If adrsTemp.State = 1 Then adrsTemp.Close
                adrsTemp.Open "SELECT deptdesc." & strKDesc & ", COUNT(*) AS TotEmp fROM Empmst INNER JOIN deptdesc ON Empmst.dept = deptdesc.dept GROUP BY deptdesc." & strKDesc & " ORDER BY deptdesc." & strKDesc & "", ConMain, adOpenStatic
                Do While Not (adrsTemp.EOF)
                    ConMain.Execute " insert into " & strRepFile & "(MndateStr,PresAbsStr) values ('" & adrsTemp.Fields(0) & "'," & adrsTemp.Fields(1) & ")", i
                    adrsTemp.MoveNext
                Loop
                                
                If adrsTemp.State = 1 Then adrsTemp.Close
                adrsTemp.Open "SELECT COUNT(*) AS Present,deptdesc." & strKDesc & " FROM Empmst INNER JOIN deptdesc ON Empmst.dept = deptdesc.dept INNER JOIN " & strTrnFile & " ON Empmst.empcode = " & strTrnFile & ".empcode GROUP BY deptdesc.dept, " & strTrnFile & ".presabs, deptdesc." & strKDesc & " HAVING (" & strTrnFile & ".presabs = 'P P ')", ConMain, adOpenStatic, adLockOptimistic
                If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                    Do While Not (adrsTemp.EOF)
                        ConMain.Execute "update " & strRepFile & " set LeaveStr=" & adrsTemp.Fields("present") & " where MndateStr='" & adrsTemp.Fields("desc") & "'"
                        adrsTemp.MoveNext
                    Loop
                End If
                If adrsTemp.State = 1 Then adrsTemp.Close
                adrsTemp.Open "SELECT COUNT(*) AS Absent,deptdesc." & strKDesc & " FROM Empmst INNER JOIN deptdesc ON Empmst.dept = deptdesc.dept INNER JOIN " & strTrnFile & " ON Empmst.empcode = " & strTrnFile & ".empcode GROUP BY deptdesc.dept, " & strTrnFile & ".presabs, deptdesc." & strKDesc & " HAVING (" & strTrnFile & ".presabs = 'A A ')", ConMain, adOpenStatic, adLockOptimistic
                If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                    Do While Not (adrsTemp.EOF)
                        ConMain.Execute "update " & strRepFile & " set PDaysStr=" & adrsTemp.Fields("Absent") & " where MndateStr='" & adrsTemp.Fields("desc") & "'"
                        adrsTemp.MoveNext
                    Loop
                End If
                If adrsTemp.State = 1 Then adrsTemp.Close
                adrsTemp.Open "SELECT COUNT(" & strTrnFile & ".presabs)AS Paid,deptdesc." & strKDesc & "  FROM " & strTrnFile & " INNER JOIN Empmst ON " & strTrnFile & ".empcode = Empmst.empcode INNER JOIN deptdesc ON Empmst.dept = deptdesc.dept GROUP BY deptdesc." & strKDesc & " , " & strTrnFile & ".presabs HAVING  (" & strTrnFile & ".presabs IN (SELECT Concat(lvcode, lvcode) FROM leavdesc  WHERE paid = 'Y' AND isitleave = 'Y')) ", ConMain, adOpenStatic, adLockOptimistic  'run for oracle only
                If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                    Do While Not (adrsTemp.EOF)
                        ConMain.Execute "update " & strRepFile & " set otStr=" & adrsTemp.Fields("Paid") & " where MndateStr='" & adrsTemp.Fields("desc") & "'"
                        adrsTemp.MoveNext
                    Loop
                End If
                If adrsTemp.State = 1 Then adrsTemp.Close
                adrsTemp.Open "SELECT COUNT(" & strTrnFile & ".presabs)AS UnPaid,deptdesc." & strKDesc & "  FROM " & strTrnFile & " INNER JOIN Empmst ON " & strTrnFile & ".empcode = Empmst.empcode INNER JOIN deptdesc ON Empmst.dept = deptdesc.dept GROUP BY deptdesc." & strKDesc & " , " & strTrnFile & ".presabs HAVING  (" & strTrnFile & ".presabs IN (SELECT Concat(lvcode, lvcode) FROM leavdesc  WHERE paid = 'N' AND isitleave = 'Y')) ", ConMain, adOpenStatic  'run for oracle only
                If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                    Do While Not (adrsTemp.EOF)
                        ConMain.Execute "update " & strRepFile & " set wrkStr=" & adrsTemp.Fields("UnPaid") & " where MndateStr='" & adrsTemp.Fields("desc") & "'"
                        adrsTemp.MoveNext
                    Loop
                End If
                

End Select
Exit Function
ERR_P:
    ShowError ("monJublientReport :: Reports")
    monJublientReport = False
End Function
Public Function peSummaryCry()
'On Error GoTo Err_P
peSummaryCry = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date

Dim STRECODE As String, strTrnFile As String, strTrnFile2 As String

dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strMon_Trn = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
'strTrnFil2 = MakeName(MonthName(Month(dtToDate)), Year(dtToDate), "trn")

End Function

' Created By
Public Function PerContAbs()
On Error GoTo ERR_P
PerContAbs = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim PresArray(31) As String
Dim DateArray(31) As Integer
Dim ArrFld As Variant
Dim arrstr As Variant
Dim bytCnt As Byte

Dim STRECODE As String, strTrnFile As String, strTrnFile2 As String
Dim valFld As String
Dim i, j, cnt
Dim strFld As String

dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strTrnFile = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strTrnFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")


If strTrnFile = strTrnFile2 Then
     
     If adrsTemp.State = 1 Then adrsTemp.Close
     Select Case InVar.strSer
        Case 1, 2 'SQL-Server,MS Access
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & "  And (Left(presabs,2) <>'P ' or Right(presabs,2) <> 'P ') " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
        Case 3 'Oracle
               adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & "  And (Lpad(presabs,2) <>'P ' or Rpad(presabs,2) <>'P ') " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
              
        
        End Select
        
Else
   
    If adrsTemp.State = 1 Then adrsTemp.Close
    Select Case InVar.strSer
    
         Case 1, 2 'SQL-server,MS Access
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode  " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Left(presabs,2) <>'P ' or Right(presabs,2) ='P ') " & strSql & _
        " Union select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile2 & ".shift,OTConf from " & strTrnFile2 & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile2 & ".Empcode  " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Left(presabs,2) <>'P ' or Right(presabs,2) <>'P ')" & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
     Case 3 ''Oracle
        
         adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode  " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Lpad(presabs,2)<>'P ' or Rpad(presabs,2) <>'P ') " & strSql & _
        " Union select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile2 & ".shift,OTConf from " & strTrnFile2 & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile2 & ".Empcode  " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And (Lpad(presabs,2) <>'P ' or Rpad(presabs,2) <>'P ')" & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
        
      End Select
        
End If
 arrstr = Array("shf", "Rem")
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
    'strDateS = monDateStr(Day(dtToDate)) 'Assigning date string
    'ReDim PresArray(31)
    Do While Not (adrsTemp.EOF) 'And dtTempDate < dtToDate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        cnt = 0
        
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    DateArray(cnt) = Day(dtTempDate)
                    PresArray(cnt) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")
                    cnt = cnt + 1
                Else
                    DateArray(cnt) = 0
                    PresArray(cnt) = ""
                    cnt = cnt + 1
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        
        valFld = "'" & STRECODE & "'"
        
strFld = "Empcode"
bytCnt = 0
For j = 0 To UBound(arrstr)
    For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
        If DateArray(i - 1) = 0 Then bytCnt = 0: Exit For
        strFld = strFld & "," & arrstr(j) & (i)
        bytCnt = bytCnt + 1
    Next i
Next j
        
         For j = 0 To 1
         ArrFld = Array(DateArray(), PresArray())
            For i = 0 To DateDiff("d", dtfromdate, dttodate) + 1
                If DateArray(i) = 0 Then Exit For
                valFld = valFld & ",'" & IIf(IsEmpty(ArrFld(j)(i)), 0, ArrFld(j)(i)) & "'"
            Next i
        Next j
        
        If bytCnt > DateDiff("d", dtfromdate, dttodate) + 1 Then
            ConMain.Execute "insert into " & strRepMfile & "(" & strFld & ") values(" & valFld & ")"
  
        End If
        strFld = ""
        bytCnt = 0
     Loop
End If
Exit Function
ERR_P:
    ShowError ("Periodic Continuous Absent :: Reports")
    PerContAbs = False
    Resume Next
End Function
' Created By  to modify by  for Monthly Memo Report Changed by
Public Function CFuncMonMemo()
On Error GoTo ERR_P
CFuncMonMemo = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim PresArray(31) As String
Dim DateArray(31) As Integer
Dim ArrFld As Variant
Dim arrstr As Variant

Dim STRECODE As String, strTrnFile As String, strDateS As String
Dim valFld As String
Dim i, j, cnt
Dim strFld As String

dtfromdate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "f"))
dttodate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))

strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")

If adrsTemp.State = 1 Then adrsTemp.Close
'*********************************************
' According Option of Report it will Take query
' contion
'********************************************
Select Case typOptIdx.bytMon
    Case 7 ' Absent Memo
       Select Case InVar.strSer
         
         Case 1, 2 ''SQL-Server,MS Access
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        "(presabs='A A ' or Left(presabs,2) ='A ' or Right(presabs,2) ='A ') and " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
        Case 3 ''Oracle
        
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        "(presabs='A A ' or Lpad(presabs,2) ='A ' or Rpad(presabs,2) ='A ') and " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
        End Select
         
         arrstr = Array("Shf", "Rem")
        
    Case 12
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " (latehrs>0 and (aflg = '0' or  aflg is  null or aflg='')) And " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        arrstr = Array("shf", "Rem")
    Case 13
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " (earlhrs>0 and (dflg = '0' or dflg is NULL or dflg=''))  And " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        arrstr = Array("Shf", "Rem")
End Select

If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
    'strDateS = monDateStr(Day(dtToDate)) 'Assigning date string
    Do While Not (adrsTemp.EOF) 'And dtTempDate < dtToDate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        cnt = 0
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    DateArray(cnt) = Day(dtTempDate)
                    Select Case typOptIdx.bytMon
                        Case 7 ''absent
                            PresArray(cnt) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")
                        Case 12 ''late
                            PresArray(cnt) = IIf(Not IsNull(adrsTemp!latehrs), Format(adrsTemp!latehrs, "0.00"), "")
                        Case 13 ''early
                            PresArray(cnt) = IIf(Not IsNull(adrsTemp!earlhrs), Format(adrsTemp!earlhrs, "0.00"), "")
                    End Select
                    cnt = cnt + 1
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        
        valFld = "'" & STRECODE & "'"
        
strFld = "Empcode"
For j = 0 To UBound(arrstr)
    For i = 1 To Day(dttodate)
        If DateArray(i - 1) = 0 Then Exit For
        strFld = strFld & "," & arrstr(j) & (i)
    Next i
Next j
        If typOptIdx.bytMon = 7 Or typOptIdx.bytMon = 12 Or typOptIdx.bytMon = 13 Then
         For j = 0 To 1
         ArrFld = Array(DateArray(), PresArray())
            For i = 0 To Day(dttodate)
                If DateArray(i) = 0 Then Exit For
                valFld = valFld & ",'" & IIf(IsEmpty(ArrFld(j)(i)), 0, ArrFld(j)(i)) & "'"
            Next i
        Next j
        End If
        If cnt > bytMode Then
            ConMain.Execute "insert into " & strRepMfile & "(" & strFld & ") values(" & valFld & ")"
            Erase PresArray: Erase DateArray
        End If
        strFld = ""
     Loop
End If

 
Exit Function
ERR_P:
    ShowError ("Monthly Performance :: Reports")
    CFuncMonMemo = False
End Function

Public Function CFuncMonShift()
On Error GoTo ERR_P
CFuncMonShift = True
Dim Query, Query1, Query2, Query3, cnt
Dim strShfFile As String, S_Str As String
Dim dtFDate As Date, dtLDate As Date, dtTempDate As Date

strShfFile = Left(typRep.strMonMth, 3) & Right(typRep.strMonYear, 2) & "Shf"
dtFDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "F"))
dtLDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "L"))
strlstdt = Day(dtLDate)

dtTempDate = dtFDate
DateStr = ""
If adrsTemp.State = 1 Then adrsTemp.Close
    Query = "insert into " & strRepMfile & " (Empcode"
    Query1 = ""
    Query2 = ""
    For cnt = 1 To 31
        Query1 = Query1 & " , Shf" & cnt
        Query2 = Query2 & " , D" & cnt
    Next
    Query3 = " select " & strShfFile & ".Empcode " & Query2 & " from " & strShfFile & "," & rpTables & _
        " where " & strShfFile & ".Empcode = empmst.Empcode " & strSql & " order by " & strShfFile & ".Empcode"
    Query = Query + Query1 + " ) " + Query3

ConMain.Execute Query


Exit Function
ERR_P:
    ShowError ("Monthly Performance :: Reports")
    CFuncMonShift = False
End Function
'Added by  for Periodic leave details 27-03
Public Function prLeaveDetails() As Boolean
On Error GoTo ERR_P
prLeaveDetails = True

Dim strLeave As String, STRECODE As String
Dim strFrLdate As String, strToFdate As String
Dim dtfromdate As String, dttodate As String
Dim strfile1 As String, strFile2 As String
Dim strFile3 As String, strFile4 As String
Dim LvArr(3) As Double, LvBalArr(3) As Double
Dim sngAvl As Double
Dim i As Integer

dtfromdate = DateCompDate(typRep.strPeriFr)
strFrLdate = DateCompStr(FdtLdt(Month(dtfromdate), Year(dtfromdate), "L"))
dttodate = DateCompDate(typRep.strPeriTo)
strToFdate = DateCompStr(FdtLdt(Month(dttodate), Year(dttodate), "F"))
strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
strFile3 = "lvbal" & Right(CStr(CInt(Year(dtfromdate))), 2)
strFile4 = "lvbal" & Right(CStr(CInt(Year(dttodate))), 2)

If adrsEmp.State = 1 Then adrsEmp.Close
    adrsEmp.Open "select distinct Empcode,empmst.cat from " & rpTables & " where " & _
    " Empcode=Empcode " & strSql & " order by Empcode ", ConMain, adOpenKeyset
    If Not (adrsEmp.EOF And adrsEmp.BOF) Then
        If adrsLeave.State = 1 Then adrsLeave.Close
            adrsLeave.Open "select * from Leavbal", ConMain, adOpenKeyset
            Do While Not adrsEmp.EOF
                STRECODE = adrsEmp("Empcode")
                For i = 1 To adrsLeave.Fields.Count - 1
                    If (UCase(adrsLeave.Fields(i).name) = "CL") Or (UCase(adrsLeave.Fields(i).name) = "PL") Or (UCase(adrsLeave.Fields(i).name) = "SL") Then
                        If adrsDept1.State = 1 Then adrsDept1.Close
                            adrsDept1.Open "select lvcode from leavdesc where cat = '" & adrsEmp("cat") & _
                            "' and lvcode = '" & adrsLeave.Fields(i).name & "'", ConMain
                            If Not (adrsDept1.EOF And adrsDept1.BOF) Then
                                strLeave = adrsLeave.Fields(i).name
                                    If strfile1 = strFile2 Then
                                        sngAvl = sngAvl + prSumAvail(STRECODE, dtfromdate, dttodate, strLeave, strfile1)
                                    Else
                                        sngAvl = sngAvl + prSumAvail(STRECODE, dtfromdate, strFrLdate, strLeave, strfile1)
                                        sngAvl = sngAvl + prSumAvail(STRECODE, strToFdate, dttodate, strLeave, strFile2)
                                    End If
                            End If
                            If UCase(strLeave) = "CL" Then
                                LvArr(1) = sngAvl
                            ElseIf UCase(strLeave) = "PL" Then
                                LvArr(2) = sngAvl
                            ElseIf UCase(strLeave) = "SL" Then
                                LvArr(3) = sngAvl
                            End If
                            sngAvl = 0
                    End If
                Next
                If adrsDept1.State = 1 Then adrsDept1.Close
                    adrsDept1.Open "select * from " & strFile4 & " where Empcode='" & STRECODE & "'", ConMain
                    If Not (adrsDept1.EOF And adrsDept1.BOF) Then
                        For i = 0 To adrsDept1.Fields.Count - 1
                            If UCase(adrsDept1.Fields(i).name) = "CL" Then
                                LvBalArr(1) = IIf(IsNull(adrsDept1.Fields(adrsDept1.Fields(i).name)), 0, adrsDept1.Fields(adrsDept1.Fields(i).name))
                            ElseIf UCase(adrsDept1.Fields(i).name) = "PL" Then
                                LvBalArr(2) = IIf(IsNull(adrsDept1.Fields(adrsDept1.Fields(i).name)), 0, adrsDept1.Fields(adrsDept1.Fields(i).name))
                            ElseIf UCase(adrsDept1.Fields(i).name) = "SL" Then
                                LvBalArr(3) = IIf(IsNull(adrsDept1.Fields(adrsDept1.Fields(i).name)), 0, adrsDept1.Fields(adrsDept1.Fields(i).name))
                            End If
                        Next
                    End If
                ConMain.Execute "insert into " & strRepFile & "(Empcode,MndateStr,PresAbsStr,LeaveStr,PDaysStr,OtStr,WrkStr)" & " values('" & STRECODE & _
                "','" & LvArr(1) & "','" & LvArr(2) & "','" & LvArr(3) & _
                "','" & LvBalArr(1) & "','" & LvBalArr(2) & "','" & LvBalArr(3) & "')"
                Erase LvArr: Erase LvBalArr
                adrsEmp.MoveNext
                If adrsEmp.EOF Then Exit Do
            Loop
    Else
        Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("00079", adrsMod), vbExclamation
        prLeaveDetails = False
    End If
Exit Function
ERR_P:
    ShowError ("prLeaveDetails :: Reports")
    prLeaveDetails = False
End Function

'Added by  for Periodic leave details 27-03
Private Function prSumAvail(ByVal STRECODE As String, ByVal strFdate As String, ByVal strLdate As String, ByVal strLvCode As String, ByVal strFileName As String) As Single
On Error GoTo ERR_P

prSumAvail = 0
If Not FindTable(strFileName) Then Exit Function
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select presabs from " & strFileName & " where empcode='" & STRECODE & "' and " & strKDate & ">=" & strDTEnc & strFdate & strDTEnc & " and " & strKDate & "<=" & strDTEnc & strLdate & strDTEnc & " order by " & strKDate, ConMain, adOpenStatic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    While Not (adrsTemp.EOF)
        If Left(adrsTemp(0), 2) = strLvCode Then
            prSumAvail = prSumAvail + 0.5
        End If
        If Right(adrsTemp(0), 2) = strLvCode Then
            prSumAvail = prSumAvail + 0.5
        End If
        adrsTemp.MoveNext
    Wend
End If
Exit Function
ERR_P:
    ShowError ("prSumAvail :: Reports")
End Function

' 18-04-09
Public Function DailyNewManpower() As Boolean
Dim dnmRs As New ADODB.Recordset
Dim dnmStr As String
Dim strDesg As String
Dim strdesg1 As String
Dim acMp As Double
Dim prnmp As Double
Dim abnmp As Double
Dim dlyNMdate As Date
Dim empgrade As String
Dim catDes As String
Dim ECode As Variant

On Error GoTo ERR_P
DailyNewManpower = True
acMp = 0
prnmp = 0
abnmp = 0

ConMain.Execute "truncate table newdlymanpower"

strMon_Trn = MakeName(MonthName(Month(DateCompDate(typRep.strDlyDate))), Year(DateCompDate(typRep.strDlyDate)), "trn")

dnmStr = "SELECT distinct empmst.designatn," & strMon_Trn & ".Empcode," & strMon_Trn & ".[Date], " & strMon_Trn & ".Shift," & strMon_Trn & ".arrtim," & strMon_Trn & ".deptim," & rpgroup1 & _
",empmst.Name,Empmst." & strKGroup & ",groupmst.grupdesc,grade.gradedesc," & strMon_Trn & ".presabs FROM " & strMon_Trn & "," & rpTables & _
" WHERE " & strMon_Trn & "." & strKDate & " = '" & Format(DateCompDate(typRep.strDlyDate), "DD-MMM-YYYY") & _
"' and (" & LeftStr(strMon_Trn & ".presabs") & " = '" & pVStar.AbsCode & _
"' or " & RightStr(strMon_Trn & ".presabs") & "='" & pVStar.AbsCode & "' or " & _
"" & LeftStr(strMon_Trn & ".presabs") & " = '" & pVStar.PrsCode & _
"' or " & RightStr(strMon_Trn & ".presabs") & "='" & pVStar.PrsCode & "' or " & _
"Left(" & strMon_Trn & ".presabs,2)='CL' OR Left(" & strMon_Trn & ".presabs,2)='SL' OR Left(" & strMon_Trn & ".presabs,2)='PL')" & _
" and " & strMon_Trn & ".Empcode = empmst.Empcode " & strSql & " order by empmst.designatn," & _
strMon_Trn & ".Empcode"

If dnmRs.State = 1 Then dnmRs.Close
dnmRs.Open dnmStr, ConMain, adOpenForwardOnly, adLockReadOnly

Do While Not dnmRs.EOF
    If strdesg1 <> "" Then
        If strDesg <> strdesg1 Or strDesg = strdesg1 Then
           ConMain.Execute "insert into Newdlymanpower(Ecode,date,gradecode,catdesc,designation,present,absent) values('" & ECode & "', '" & dlyNMdate & "','" & empgrade & "', '" & catDes & "','" & strdesg1 & "'," _
           & " " & prnmp & " ," & abnmp & ")"
           'acMp = 0
           prnmp = 0
           abnmp = 0
         End If
    End If

   If strDesg = strdesg1 Or dnmRs.Fields("designatn") = strDesg Then
    If dnmRs.Fields("arrtim") <> 0 And dnmRs.Fields("deptim") <> 0 Or dnmRs.Fields("cat") = "006" Then
        If dnmRs.Fields("presabs") = pVStar.PrsCode & pVStar.PrsCode Then
            prnmp = prnmp + 1
        ElseIf dnmRs.Fields("presabs") = pVStar.PrsCode & pVStar.AbsCode Then
            prnmp = prnmp + 0.5
        ElseIf dnmRs.Fields("presabs") = pVStar.AbsCode & pVStar.PrsCode Then
            prnmp = prnmp + 0.5
        End If
    Else
        If dnmRs.Fields("presabs") = pVStar.AbsCode & pVStar.AbsCode Then
            abnmp = abnmp + 1
        ElseIf Left(dnmRs.Fields("presabs"), 2) = "CL" Then
            abnmp = abnmp + 1
        ElseIf Left(dnmRs.Fields("presabs"), 2) = "SL" Then
            abnmp = abnmp + 1
        ElseIf Left(dnmRs.Fields("presaBs"), 2) = "PL" Then
            abnmp = abnmp + 1
        End If
    End If
  End If
        
    strdesg1 = dnmRs.Fields("designatn")
    empgrade = dnmRs.Fields("gradecode")
    catDes = dnmRs.Fields("gradedesc")
    dlyNMdate = dnmRs.Fields("Date")
    ECode = dnmRs.Fields("Empcode")
    
    'acMp = acMp + 1
    
    

  strDesg = dnmRs.Fields("designatn")
      ConMain.Execute "insert into Newdlymanpower(Ecode,date,gradecode,catdesc,designation,present,absent) values('" & ECode & "', '" & dlyNMdate & "','" & empgrade & "','" & catDes & "','" & strdesg1 & "'," _
    & " " & prnmp & " ," & abnmp & ")"
'When last record going to be inserted into table
  On Error Resume Next
  dnmRs.MoveNext
Loop
        
'conmain.Execute "insert into Newdlymanpower(Ecode,date,gradecode,catdesc,designation,present,absent) values('" & Ecode & "', '" & dlyNMdate & "','" & empgrade & "','" & catDes & "','" & strdesg1 & "'," _
'& " " & prnmp & " ," & abnmp & ")"
abnmp = 0
prnmp = 0

Exit Function
ERR_P:
    ShowError ("DailyNewManpower :: Reports")
    DailyNewManpower = False
End Function

' 18-04-09
Public Function NewAttendance()
On Error GoTo ERR_P
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim strfile1 As String, strFile2 As String
Dim rsAttend As New ADODB.Recordset
Dim sqlAtt As String
Dim lvstr, STRECODE As String
Dim lvA, LvP, lvCL, lvPL, lvSL, lvHL As Single
NewAttendance = True

lvA = 0: LvP = 0: lvCL = 0: lvSL = 0: lvPL = 0: lvHL = 0

dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)

strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

ConMain.Execute "truncate table PNewAtt"

sqlAtt = "select " & strfile1 & ".Empcode," & strKDate & "," & _
"presabs from " & strfile1 & "," & rpTables & " where " & _
strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & strSql & " union select " & strFile2 & ".Empcode," & strKDate & "," & _
"presabs from " & strFile2 & "," & rpTables & " where " & strFile2 & ".Empcode = empmst.Empcode " & _
"and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " " & strSql

Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        sqlAtt = sqlAtt & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        sqlAtt = sqlAtt & " order by Empcode," & strKDate
End Select

If rsAttend.State = 1 Then rsAttend.Close
rsAttend.Open sqlAtt, ConMain, adOpenForwardOnly, adLockReadOnly

If Not (rsAttend.BOF And rsAttend.EOF) Then

    If Not (IsNull(rsAttend(0)) Or IsEmpty(rsAttend(0))) Then
       rsAttend.MoveFirst

        Do While Not (rsAttend.EOF) And dtfromdate <= dttodate
            STRECODE = rsAttend!Empcode
            dtfromdate = dtFirstDate
  
            Do While dtfromdate <= dttodate
                If rsAttend.EOF Then Exit Do
                If rsAttend!Empcode = STRECODE Then
                    If rsAttend!Date = dtfromdate Then
                        
                        If rsAttend!presabs = pVStar.HlsCode & pVStar.HlsCode Then
                            lvHL = lvHL + 1
                        ElseIf rsAttend!presabs = "CLCL" Or rsAttend!presabs = "CL" Then
                            lvCL = lvCL + 1
                        ElseIf rsAttend!presabs = "PLPL" Or rsAttend!presabs = "PL" Then
                            lvPL = lvPL + 1
                        ElseIf rsAttend!presabs = "SLSL" Or rsAttend!presabs = "SL" Then
                            lvSL = lvSL + 1
                        ElseIf rsAttend!presabs = pVStar.AbsCode & pVStar.AbsCode Then
                            lvA = lvA + 1
                        ElseIf rsAttend!presabs = pVStar.PrsCode & pVStar.PrsCode Then
                            LvP = LvP + 1
                        End If
                   
                        If dtfromdate = rsAttend!Date Then rsAttend.MoveNext
'                            dtfromdate = DateAdd("d", 1, dtfromdate) ' 23-04-09
                     End If
                            dtfromdate = DateAdd("d", 1, dtfromdate)
                Else
                   Exit Do
                End If
            Loop
            
            ConMain.Execute "insert into PNewAtt (Empcode,P,A,CL,SL,PL,HL)values('" & STRECODE & "'," & LvP & "," & lvA & "," & lvCL & "," & lvSL & "," & lvPL & "," & lvHL & ")"
            lvA = 0: LvP = 0: lvCL = 0: lvSL = 0: lvPL = 0: lvHL = 0
            dtfromdate = dtFirstDate
            If rsAttend.EOF = False Then rsAttend.MoveNext ' 27-04
        Loop
    End If
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    NewAttendance = False
End If
Exit Function
ERR_P:
    ShowError ("Periodic NewAttendance :: Reports")
    NewAttendance = False
    'Resume Next
End Function


'  18-04-09
Public Function NewMonAttendance() As Boolean

End Function

Public Function yrSummery() As Boolean   ' for Glenmark
On Error GoTo ERR_P
yrSummery = True
Dim strYrMan As String, strLVselect As String, Present As String, strCurMon As String, strCurYear As String
Dim strFdate As String, strLdate As String, STRECODE As String, strFileName As String
Dim i As Integer, k As Integer, j As Integer, TotValStr As String, PresCnt As Integer
Dim PdValStr As String, MLValStr As String, LWValStr As String, strGY As String, ALValStr As String
Dim SLValStr As String, CLValStr As String, WOValStr As String, AbStr As String, StrTotArr As String, WO1ValStr As String, PLValStr As String
Dim TotVal As Double

Present = "P ": PresCnt = 1
strFdate = FdtLdt(CByte(pVStar.Yearstart), typRep.strYear, "f")
strLdate = FdtLdt(CByte(pVStar.Yearstart) - 1, IIf(pVStar.Yearstart = "1", _
typRep.strYear, typRep.strYear + 1), "l")
strCurMon = strFdate
strCurYear = Year(DateCompDate(strFdate))
If adrsEmp.State = 1 Then adrsEmp.Close
strFileName = "lvtrn" & Right(strYearFrom(strFdate), 2)
adrsEmp.Open "select DISTINCT empmst.EMPCODE from " & strFileName & "," & rpTables & " where " & _
"joindate <= " & strDTEnc & DateCompStr(strLdate) & strDTEnc & " and " & strFileName & ".empcode =empmst.empcode " & strSql & _
" ORDER BY empmst.EMPCODE", ConMain, adOpenStatic
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "select distinct lvcode from leavdesc where  isitleave ='y' and paid='y' and lvcode not in ('CL','PL','AL','ML','SL','OD','" & pVStar.AbsCode & "','" & pVStar.PrsCode & "','" & pVStar.WosCode & "','" & pVStar.HlsCode & "')", ConMain, adOpenStatic
If Not (adrsLeave.BOF And adrsLeave.EOF) Then
    For i = 1 To adrsLeave.RecordCount
        If i <> adrsLeave.RecordCount Then
            strLVselect = strLVselect & adrsLeave(0) & ","
        Else
            strLVselect = strLVselect & adrsLeave(0)
        End If
        adrsLeave.MoveNext
    Next i
End If
Dim strArr() As Variant
Dim FldArr(0 To 8) As Integer
Dim strArrTmp() As String
Dim TotArr(1 To 12) As Double
If FieldExists(strFileName, "OD") Then
    Present = Present & ",OD"
    PresCnt = PresCnt + 1
End If
strArrTmp = Split(strLVselect, ",")
strArr = Array(Present, "ML", strLVselect, "SL", "CL", "HL", "A ", "PL", "AL")
For k = 0 To UBound(strArr)
    If FieldExists(strFileName, strArr(k)) Then
        FldArr(k) = 1
    Else
        FldArr(k) = 0
    End If
Next
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp!Empcode
        strCurMon = strFdate
        strCurYear = Year(DateCompDate(strFdate))
        For k = 0 To UBound(strArr)
            If FldArr(k) = 1 Or k = 2 Or k = 0 Then
                Dim LvPd As Double
                For i = 1 To 12
                    strFdate = FdtLdt(Month(DateCompDate(strCurMon)), strCurYear, "f")
                    strLdate = FdtLdt(Month(DateCompDate(strCurMon)), strCurYear, "l")
                    If adrsTemp.State = 1 Then adrsTemp.Close
                    adrsTemp.Open "select " & Year(strLdate) & "," & strArr(k) & _
                    " from " & strFileName & " where Empcode=" & "'" & STRECODE & "'" & " and lst_date=" & _
                    strDTEnc & DateCompStr(strLdate) & strDTEnc, ConMain, adOpenStatic
                    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                        If k = 2 Then
                            LvPd = 0
                            For j = 0 To UBound(strArrTmp)
                                LvPd = LvPd + IIf(Not IsNull(adrsTemp(strArrTmp(j))) And adrsTemp(strArrTmp(j)) > 0, adrsTemp(strArrTmp(j)), 0)
                            Next
                            strYrMan = strYrMan & IIf(LvPd > 0, Format(LvPd, "0.00") & Spaces(Len(Format(LvPd, "0.00"))), Spaces(0))
                            TotArr(i) = TotArr(i) + IIf(LvPd > 0, LvPd, 0)
                            TotVal = TotVal + IIf(LvPd > 0, LvPd, 0)
                        ElseIf k = 0 Then
                            LvPd = 0
                            For j = 1 To PresCnt
                                LvPd = LvPd + IIf(Not IsNull(adrsTemp(j)) And adrsTemp(j) > 0, adrsTemp(j), 0)
                            Next
                            strYrMan = strYrMan & IIf(LvPd > 0, Format(LvPd, "0.00") & Spaces(Len(Format(LvPd, "0.00"))), Spaces(0))
                            TotArr(i) = TotArr(i) + IIf(LvPd > 0, LvPd, 0)
                            TotVal = TotVal + IIf(LvPd > 0, LvPd, 0)
                        Else
                            strYrMan = strYrMan & IIf(Not IsNull(adrsTemp(1)) And adrsTemp(1) > 0, _
                            Format(adrsTemp(1), "0.00") & Spaces(Len(Format(adrsTemp(1), "0.00"))), Spaces(0))
                            TotArr(i) = TotArr(i) + IIf(Not IsNull(adrsTemp(1)) And adrsTemp(1) > 0, adrsTemp(1), 0)
                            TotVal = TotVal + IIf(Not IsNull(adrsTemp(1)) And adrsTemp(1) > 0, adrsTemp(1), 0)
                        End If
                    Else
                        strYrMan = strYrMan & Spaces(0)
                        TotArr(i) = TotArr(i) + 0
                        TotVal = TotVal + 0
                    End If
                    strCurMon = CStr(DateAdd("m", 1, DateCompDate(strCurMon)))
                    If i <> 12 Then
                        If Month(DateCompDate(strCurMon)) < Month(DateCompDate(strLdate)) Then
                            strCurYear = Year(DateCompDate(strFdate)) + 1
                        End If
                    Else
                        strFdate = FdtLdt(CByte(pVStar.Yearstart), typRep.strYear, "f")
                        strLdate = FdtLdt(CByte(pVStar.Yearstart) - 1, IIf(pVStar.Yearstart = "1", _
                        typRep.strYear, typRep.strYear + 1), "l")
                        strCurMon = strFdate
                        strCurYear = Year(DateCompDate(strFdate))
                    End If
                Next i
            Else
                For i = 1 To 12
                    strYrMan = strYrMan & Spaces(0)
                    TotArr(i) = TotArr(i) + 0
                    TotVal = TotVal + 0
                Next
            End If
            strYrMan = strYrMan & Format(CStr(TotVal), "0.00")
            Select Case strArr(k)
                Case Present: PdValStr = strYrMan
                Case "ML": MLValStr = strYrMan
                Case strLVselect: LWValStr = strYrMan
                Case "SL": SLValStr = strYrMan
                Case "CL": CLValStr = strYrMan
                Case "HL": WOValStr = strYrMan
                Case "A ": AbStr = strYrMan
                Case "PL": PLValStr = strYrMan
                Case "AL": ALValStr = strYrMan
                Case Else
            End Select
            TotValStr = TotValStr & TotVal & ","
            strYrMan = "": TotVal = 0
        Next k
        For i = 1 To 12
            'StrTotArr = StrTotArr & IIf(TotArr(i) > 0, Format(TotArr(i), "0.00") & Spaces(Len(Format(TotArr(i), "0.00"))), Spaces(0))
            StrTotArr = StrTotArr & Format(TotArr(i), "0.00") & Spaces(Len(Format(TotArr(i), "0.00")))
        Next
        
        strGY = "insert into " & strRepFile & " (Empcode,Yr,PStr,MLStr,StrikeStr,LWStr,LayOffStr,SLStr,CLStr,PdHLStr,AbStr,MnthTot) " & _
        " values('" & STRECODE & "','" & strCurYear & "','" & PdValStr & "','" & MLValStr & "','" & PLValStr & "','" & LWValStr & "','" & ALValStr & "','" & SLValStr & "','" & CLValStr & "','" & WOValStr & "','" & AbStr & "','" & StrTotArr & "')"
        TotVal = 0: StrTotArr = ""
        If Trim(strGY) <> "" Then ConMain.Execute strGY
        strYrMan = "": PdValStr = "": MLValStr = "": LWValStr = "": SLValStr = ""
        CLValStr = "": PLValStr = "": WOValStr = "": AbStr = "": TotValStr = "": Erase TotArr
        strCurYear = Val(typRep.strYear)
        strFdate = FdtLdt(Month(DateCompDate(strCurMon)), strCurYear, "f")
        adrsEmp.MoveNext
        If bytBackEnd = 2 Then Sleep (100)
    Loop
End If
Erase FldArr
Exit Function
ERR_P:
    ShowError ("Yearly summery :: Reports")
    yrSummery = False
    'Resume Next
End Function

Public Function yrAbsentism()   'Added by  18-11
On Error GoTo ERR_P
yrAbsentism = True
Dim adrsLvCD As New ADODB.Recordset
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim strqry As String, STRECODE As String, strName As String, strTrnFile As String
Dim i, j As Integer
Dim strFld As String, valFld As String, Mnth As String, absent As String
Dim strFdate As String, strLdate As String, strCurYear As String, strCurMon As String

strFld = "Empcode,trcd,ystr,YValStr,[Counter]"
strFdate = FdtLdt(CByte(pVStar.Yearstart), typRep.strYear, "F")
strLdate = FdtLdt(CByte(pVStar.Yearstart) - 1, IIf(pVStar.Yearstart = "1", typRep.strYear, typRep.strYear + 1), "L")
    
If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select DISTINCT EMPCODE,NAME from " & rpTables & " where joindate <= " & _
strDTEnc & DateCompStr(strLdate) & strDTEnc & " " & strSql & " ORDER BY EMPCODE", ConMain, adOpenStatic
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp!Empcode
        strName = adrsEmp!name
        strCurMon = strFdate
        strCurYear = strYearFrom(strFdate)
        For i = 1 To 12
            strTrnFile = MakeName(MonthName(Month(strCurMon)), strCurYear, "trn")
            If FindTable(strTrnFile) Then
                dtfromdate = DateCompDate(FdtLdt(MonthNumber(MonthName(Month(strCurMon))), typRep.strYear, "f"))
                dttodate = DateCompDate(FdtLdt(MonthNumber(MonthName(Month(strCurMon))), typRep.strYear, "l"))
                strqry = "select empmst.Empcode," & strKDate & ",presabs from " & strTrnFile & "," & rpTables & " where " & _
                        " empmst.Empcode = " & strTrnFile & ".Empcode and (presabs=" & "'" & ReplicateVal(pVStar.AbsCode, 2) & "'" & _
                        " or " & LeftStr("presabs") & "=" & "'" & pVStar.AbsCode & "'" & " or " & RightStr("presabs") & "=" & _
                        "'" & pVStar.AbsCode & "'" & ")" & strSql & " and " & strTrnFile & ".Empcode='" & STRECODE & "'" & " order by empmst.Empcode," & strKDate & ""
                If adrsTemp.State = 1 Then adrsTemp.Close
                adrsTemp.Open strqry, ConMain, adOpenStatic
                If Not (adrsTemp.BOF And adrsTemp.EOF) Then
                    Mnth = UCase(Left(strTrnFile, 3)) & "-" & Mid(strTrnFile, 4, 2)
                    dtTempDate = dtfromdate
                    Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
                        If adrsTemp!Date = dtTempDate Then
                            absent = absent & Format(Day(dtTempDate), "00") & "-" & adrsTemp!presabs & " "
                        End If
                        If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
                        dtTempDate = DateAdd("d", 1, dtTempDate)
                    Loop
                    valFld = "'" & STRECODE & "','" & strName & "','" & Mnth & "','" & absent & "'," & i
                    ConMain.Execute "insert into " & strRepFile & " (" & strFld & ") values(" & valFld & ")"
                    valFld = "": absent = ""
                End If
            End If
            strCurMon = CStr(DateAdd("m", 1, DateCompDate(strCurMon)))
            If Month(DateCompDate(strCurMon)) <= Month(DateCompDate(strLdate)) And pVStar.Yearstart <> 1 Then
                strCurYear = Year(DateCompDate(strFdate)) + 1
            End If
        Next
        adrsEmp.MoveNext
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Yearly Absentism :: Reports")
    yrAbsentism = False
End Function

Public Function peAttendance()  ' 14-01
On Error GoTo ERR_P
peAttendance = True
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim strfile1 As String, strFile2 As String, DTESTR As String, STRECODE As String
Dim p_str As String, strGP As String
Dim SecShft As Single
Dim valFld As String

dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)
strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
DTESTR = ""
Do While dtfromdate <= dttodate
    If dtfromdate = dttodate Then
        DTESTR = DTESTR & Day(dtfromdate)
    ElseIf dtfromdate <> dttodate Then
        DTESTR = DTESTR & Day(dtfromdate) & Spaces(Len(Trim(str(Day(dtfromdate)))))
    End If
    dtfromdate = DateAdd("d", 1, dtfromdate)
Loop
If strfile1 = strFile2 Then
    strGP = "select " & strfile1 & ".Empcode," & strfile1 & ".shift,arrtim,deptim,wrkhrs," & strKDate & ",presabs from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
    strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " " & strSql
Else
    strGP = "select " & strfile1 & ".Empcode," & strfile1 & ".shift,arrtim,deptim,wrkhrs," & strKDate & ",presabs from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & Format(DateCompDate(typRep.strPeriFr), "dd/mmm/yy") & _
    strDTEnc & " " & strSql & _
    " union select " & strFile2 & ".Empcode," & strFile2 & ".shift,arrtim,deptim,wrkhrs," & strKDate & ",presabs from " & strFile2 & "," & rpTables & " where " & _
    strFile2 & ".Empcode = empmst.Empcode and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "dd/mmm/yy") & strDTEnc & " " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    adrsTemp.MoveFirst
    Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
        STRECODE = adrsTemp!Empcode
        dtfromdate = dtFirstDate
        p_str = ""
        Do While dtfromdate <= dttodate
            If adrsTemp.EOF Then Exit Do
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtfromdate Then
                    p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                    If UCase(Trim(adrsTemp!Shift)) = "B" And (adrsTemp!arrtim > 0 Or adrsTemp!deptim > 0) Then
                            SecShft = SecShft + 1
                        End If
                ElseIf adrsTemp!Date <> dtfromdate Then
                    p_str = p_str & Spaces(0)
                End If
            Else
                Exit Do
            End If
            If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
            dtfromdate = DateAdd("d", 1, dtfromdate)
        Loop
        If Trim(p_str) = "" Then
        Else
            valFld = valFld & IIf(SecShft <> 0, Format(CStr(SecShft), "0.00"), "0.00") & IIf(Len(CStr(SecShft)) >= 6, "", Spaces(Len(Format(CStr(SecShft), "0.00"))))
            ConMain.Execute "insert into " & strRepFile & "" & _
            "(Empcode," & strKDate & ",presabsstr,arrstr)  values('" & STRECODE & _
            "','" & DTESTR & "','" & p_str & "','" & valFld & "')"
        End If
        SecShft = 0: valFld = "": dtfromdate = dtFirstDate
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Periodic Attendance :: Reports")
    peAttendance = False
End Function

Public Function MusterData() As Boolean ' 26-02
On Error GoTo ERR_P
MusterData = True
Dim FDt1 As Date, ToDt1 As Date, dtTempDate As Date
Dim strTrnFile1 As String, strTrnFile2 As String, strLv As String
Dim PSum As Single, PLSum As Single, SLSum As Single, CLSum As Single
Dim SType As String, Shift As String, STime As String
Dim STRECODE As String, p_str As String, S_Str As String

FDt1 = typRep.strPeriFr
ToDt1 = typRep.strPeriTo
strlstdt = Day(ToDt1)
strTrnFile1 = MakeName(MonthName(Month(typRep.strPeriFr)), Year(typRep.strPeriFr), "trn")
strTrnFile2 = MakeName(MonthName(Month(typRep.strPeriTo)), Year(typRep.strPeriTo), "trn")

If adrsEmp.State = 1 Then adrsEmp.Close
If strTrnFile1 = strTrnFile2 Then
    adrsEmp.Open "select distinct empmst.Empcode,empmst.styp,empmst.f_shf,empmst.scode,presabs," & strKDate & " from " & strTrnFile1 & ",empmst,deptdesc where " & _
    " empmst.Empcode = " & strTrnFile1 & ".Empcode and " & strKDate & ">=" & strDTEnc & Format(typRep.strPeriFr, "dd/mmm/yy") & strDTEnc & " and " & strKDate & "<=" & strDTEnc & Format(typRep.strPeriTo, "dd/mmm/yy") & strDTEnc & " " & strSql & _
    " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
Else
    adrsEmp.Open "select distinct empmst.Empcode,empmst.styp,empmst.f_shf,empmst.scode,presabs," & strKDate & " from " & strTrnFile1 & ",empmst,deptdesc where " & _
    " empmst.Empcode = " & strTrnFile1 & ".Empcode and " & strKDate & ">=" & strDTEnc & Format(typRep.strPeriFr, "dd/mmm/yy") & strDTEnc & " " & strSql & _
    " union select distinct empmst.Empcode,empmst.styp,empmst.f_shf,empmst.scode,presabs," & strKDate & " from " & strTrnFile2 & ",empmst,deptdesc where " & _
    " empmst.Empcode = " & strTrnFile2 & ".Empcode and " & strKDate & "<=" & strDTEnc & Format(typRep.strPeriTo, "dd/mmm/yy") & strDTEnc & " " & strSql & _
    " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
End If
If Not (adrsEmp.BOF And adrsEmp.EOF) Then
    dtTempDate = FDt1
    Do While Not (adrsEmp.EOF)
        STRECODE = adrsEmp!Empcode
        SType = adrsEmp!styp
        If SType = "F" Then
            Shift = adrsEmp!f_shf
        Else
            Shift = adrsEmp!scode
        End If
        dtTempDate = FDt1
        p_str = ""
        Do While dtTempDate <= ToDt1 And Not (adrsEmp.EOF)
            If adrsEmp!Empcode = STRECODE Then
                If adrsEmp!Date = dtTempDate Then
                    p_str = p_str & IIf(Not IsNull(adrsEmp!presabs), adrsEmp!presabs & Spaces(Len(adrsEmp!presabs)), "      ")
                    If Not IsNull(adrsEmp!presabs) Then
                        Select Case Left(adrsEmp!presabs, 2)
                            Case "P ": PSum = PSum + 0.5
                            Case "PL": PLSum = PLSum + 0.5
                            Case "SL": SLSum = SLSum + 0.5
                            Case "CL": CLSum = CLSum + 0.5
                        End Select
                        Select Case Right(adrsEmp!presabs, 2)
                            Case "P ": PSum = PSum + 0.5
                            Case "PL": PLSum = PLSum + 0.5
                            Case "SL": SLSum = SLSum + 0.5
                            Case "CL": CLSum = CLSum + 0.5
                        End Select
                    End If
                ElseIf adrsEmp!Date <> dtTempDate Then
                    p_str = p_str & "    " & Spaces(0)
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsEmp!Date Then adrsEmp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
        Loop
        If SType = "F" Then
            If adrsTemp.State = 1 Then adrsTemp.Close
            adrsTemp.Open "select shf_in,shf_out from instshft where shift='" & Shift & "'", ConMain, adOpenStatic
            STime = STime & Format(adrsTemp!SHF_IN, "00.00") & Space(1) & Format(adrsTemp!shf_out, "00.00") & Space(1)
        Else
            If adrsTemp.State = 1 Then adrsTemp.Close
            adrsTemp.Open "select pattern from ro_shift where scode='" & Shift & "'", ConMain, adOpenStatic
            Dim arrShft() As String, arrTShft() As String
            Dim i As Integer, j As Integer
            arrTShft = Split(Mid(adrsTemp!Pattern, 1, Len(adrsTemp!Pattern) - 1), ".")
            Shift = ""
            For i = 0 To UBound(arrTShft)
                If arrTShft(i) <> "WO" And arrTShft(i) <> "HL" Then
                    Shift = Shift & "'" & arrTShft(i) & "'" & ","
                    ReDim Preserve arrShft(j)
                    arrShft(j) = arrTShft(i)
                    j = j + 1
                End If
            Next
            Erase arrTShft
            Shift = Mid(Shift, 1, Len(Shift) - 1)
            If adrsTemp.State = 1 Then adrsTemp.Close
            adrsTemp.Open "select shift,shf_in,shf_out from instshft where shift IN (" & Shift & ")", ConMain, adOpenStatic
            For i = 0 To UBound(arrShft)
                adrsTemp.MoveFirst
                Do While Not (adrsTemp.EOF)
                    If arrShft(i) = adrsTemp!Shift Then
                        STime = STime & Format(IIf(Val(adrsTemp!SHF_IN) > 23.59, Val(adrsTemp!SHF_IN) - 24, adrsTemp!SHF_IN), "00.00") & Space(1) & Format(IIf(Val(adrsTemp!shf_out) > 23.59, Val(adrsTemp!shf_out) - 24, adrsTemp!shf_out), "00.00") & Space(1)
                        Exit Do
                    Else
                        adrsTemp.MoveNext
                    End If
                Loop
            Next
        End If
        strLv = PSum & Space(1) & PLSum & Space(1) & SLSum & Space(1) & CLSum
        If Trim(p_str) = "" Then
        Else
            ConMain.Execute "insert into " & strRepFile & "(Empcode," & strKDate & ",presabsstr,shfstr) " & _
            " values(" & "'" & STRECODE & "'" & "," & "'" & strLv & "'" & "," & "'" & p_str & "','" & Trim(STime) & "')"
        End If
        SType = "": STime = "": Shift = "": Erase arrShft: PSum = 0: PLSum = 0: SLSum = 0: CLSum = 0: strLv = ""
    Loop
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbExclamation
    MusterData = False
End If
Exit Function
ERR_P:
    ShowError ("MUSTER DATA::")
    MusterData = False
End Function

Public Function WKFormJ() As Boolean    ' 14-04
On Error GoTo ERR_P
WKFormJ = True
Dim wsum As Single, osum As Single
Dim A_Str As String, D_Str As String, W_Str As String, O_Str As String, Dt_Str As String
Dim RstIn As String, RstOut As String, strGP As String
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim STRECODE As String, strfile1 As String, strFile2 As String
Dim i  As Integer
i = 1
dtFirstDate = DateCompDate(typRep.strWkDate)
dtfromdate = DateCompDate(typRep.strWkDate)
dttodate = DateCompDate(typRep.strWkDate) + 6
strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
If strfile1 = strFile2 Then
    strGP = "SELECT " & strfile1 & ".Empcode," & strKDate & ",name,sex,birth_dt,arrtim, deptim, ovtim,OTConf,wrkhrs, hdend, hdstart, rst_in, rst_out, rst_in_2, rst_out_2, rst_in_3, rst_out_3 From " & rpTables & ", " & _
    strfile1 & ", instshft Where " & strfile1 & "." & strKDate & ">=" & strDTEnc & Format(DateCompDate(dtFirstDate), "dd/MMM/yy") & _
    strDTEnc & " and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(dttodate), "dd/MMM/yy") & strDTEnc & " and Empmst.Empcode = " & strfile1 & ".Empcode And " & strfile1 & ".Shift = instshft.Shift " & strSql
Else
    strGP = "SELECT " & strfile1 & ".Empcode," & strKDate & ",name,sex,birth_dt,arrtim, deptim, ovtim,OTConf,wrkhrs, hdend, hdstart, rst_in, rst_out, rst_in_2, rst_out_2, rst_in_3, rst_out_3 From " & rpTables & ", " & _
    strfile1 & ", instshft Where " & strfile1 & "." & strKDate & ">=" & strDTEnc & Format(DateCompDate(dtFirstDate), "dd/MMM/yy") & _
    strDTEnc & " and Empmst.Empcode = " & strfile1 & ".Empcode And " & strfile1 & ".Shift = instshft.Shift " & _
    strSql & " union SELECT " & strFile2 & ".Empcode," & strKDate & ",name,sex,birth_dt,arrtim, deptim, ovtim,OTConf,wrkhrs, hdend, hdstart, rst_in, rst_out, rst_in_2, rst_out_2, rst_in_3, rst_out_3 From " & rpTables & ", " & _
    strFile2 & ", instshft Where " & strFile2 & "." & strKDate & "<=" & strDTEnc & Format(DateCompDate(dttodate), "dd/MMM/yy") & strDTEnc & " and Empmst.Empcode = " & strFile2 & ".Empcode And " & strFile2 & ".Shift = instshft.Shift " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
i = 1
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then    'adrsTemp
    If Not (IsNull(adrsTemp(0)) Or IsEmpty(adrsTemp(0))) Then
        adrsTemp.MoveFirst
        Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
            STRECODE = adrsTemp!Empcode
            dtfromdate = dtFirstDate
            Do While dtfromdate <= dttodate
                If adrsTemp.EOF Then Exit Do
                If adrsTemp!Empcode = STRECODE Then
                    Dt_Str = Dt_Str & Format(DateCompDate(adrsTemp!Date), "dd/MMM/yy") & " "
                    If adrsTemp!Date = dtfromdate Then
                        A_Str = A_Str & IIf(Not IsNull(adrsTemp!arrtim) And adrsTemp!arrtim > 0, _
                        Spaces(Len(Format(adrsTemp!arrtim, "0.00"))) & Format(adrsTemp!arrtim, "0.00"), "  0.00")
                                                        
                        D_Str = D_Str & IIf(Not IsNull(adrsTemp!deptim) And adrsTemp!deptim > 0, _
                        Spaces(Len(Format(adrsTemp!deptim, "0.00"))) & Format(adrsTemp!deptim, "0.00"), "  0.00")
                                                                                                                                        
                        W_Str = W_Str & IIf(Not IsNull(adrsTemp!wrkHrs) And adrsTemp!wrkHrs > 0, _
                        Spaces(Len(Format(adrsTemp!wrkHrs, "0.00"))) & Format(adrsTemp!wrkHrs, "0.00"), "  0.00")
                        
                        RstIn = RstIn & IIf(Not IsNull(adrsTemp!hdend) And adrsTemp!hdend > 0, _
                        Spaces(Len(Format(adrsTemp!hdend, "0.00"))) & Format(adrsTemp!hdend, "0.00"), "  0.00")
                        
                        RstOut = RstOut & IIf(Not IsNull(adrsTemp!hdstart) And adrsTemp!hdstart > 0, _
                        Spaces(Len(Format(adrsTemp!hdstart, "0.00"))) & Format(adrsTemp!hdstart, "0.00"), "  0.00")
                        
                        'If adrsTemp("OTConf") = "Y" Then     ''if authorized OT then only Calculate and show
                            O_Str = O_Str & IIf(Not IsNull(adrsTemp!ovtim) And adrsTemp!ovtim > 0, _
                            Spaces(Len(Format(adrsTemp!ovtim, "0.00"))) & Format(adrsTemp!ovtim, "0.00"), "  0.00")
                            osum = TimAdd(IIf(IsNull(osum), 0, osum), IIf(IsNull(adrsTemp!ovtim), 0, adrsTemp!ovtim))
                         'Else
                         '   O_Str = O_Str & "  0.00"
                        'End If
                        wsum = TimAdd(IIf(IsNull(wsum), 0, wsum), IIf(IsNull(adrsTemp!wrkHrs), 0, adrsTemp!wrkHrs))
                    ElseIf adrsTemp!Date <> dtfromdate Then
                        A_Str = A_Str & "  0.00"
                        D_Str = D_Str & "  0.00"
                        W_Str = W_Str & "  0.00"
                        O_Str = O_Str & "  0.00"
                        RstIn = RstIn & "  0.00"
                        RstOut = RstOut & "  0.00"
                    End If
                Else
                    Exit Do
                End If
                If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
                dtfromdate = DateAdd("d", 1, dtfromdate)
            i = i + 1
            If i > 7 Then Exit Do
            Loop 'END OF DATE LOOP
            wsum = IIf(wsum = 0, "0.00", Format(wsum, "0.00"))
            osum = IIf(osum = 0, "0.00", Format(osum, "0.00"))

            ConMain.Execute "insert into " & strRepFile & "" & _
            "(Empcode,Arrstr,DepStr,WorkStr,LateStr,EarlStr,OTStr," & strKDate & ",sumwork,sumOT) values ('" & _
            STRECODE & "','" & A_Str & "','" & D_Str & "','" & W_Str & "','" & RstIn & "','" & RstOut & "','" & O_Str & "','" & Dt_Str & "'," & wsum & "," & osum & ")"
            wsum = 0: osum = 0
            A_Str = "": D_Str = "": W_Str = "": O_Str = "": RstIn = "": RstOut = "": Dt_Str = ""
            i = 1
            dtfromdate = dtFirstDate
        Loop 'END OF EMPLOYEE LOOP
    End If
Else
    Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("00079", adrsMod), vbInformation
    WKFormJ = False
End If 'adrsTemp.eof
Exit Function
ERR_P:
    ShowError ("Periodic Performance Overtime :: Reports")
    WKFormJ = False
''    Resume Next
End Function

Public Function Attendance()               '
On Error GoTo ERR_P
Attendance = True
Dim dtfromdate As Date, dttodate As Date, dtFirstDate As Date
Dim strfile1 As String, strFile2 As String, DTESTR As String, STRECODE As String
Dim p_str As String, strGP As String
Dim SecShft As Single
Dim valFld As String
Dim sngCL As Single, sngSL As Single, sngpp As Single, sngAA As Single, sngOD As Single, snglv As Single, sngHL As Single, sngWO As Single, paid As Single
Dim sngML As Single, sngCO As Single, sngPL As Single
dtFirstDate = DateCompDate(typRep.strPeriFr)
dtfromdate = DateCompDate(typRep.strPeriFr)
dttodate = DateCompDate(typRep.strPeriTo)
strfile1 = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")
DTESTR = ""
Do While dtfromdate <= dttodate
    If dtfromdate = dttodate Then
        DTESTR = DTESTR & Day(dtfromdate)
    ElseIf dtfromdate <> dttodate Then
        DTESTR = DTESTR & Day(dtfromdate) & Spaces(Len(Trim(str(Day(dtfromdate)))))
    End If
    dtfromdate = DateAdd("d", 1, dtfromdate)
Loop
If strfile1 = strFile2 Then
    strGP = "select " & strfile1 & ".Empcode," & strfile1 & ".shift,arrtim,deptim,wrkhrs," & strKDate & ",presabs from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
    strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " " & strSql
Else
    strGP = "select " & strfile1 & ".Empcode," & strfile1 & ".shift,arrtim,deptim,wrkhrs," & strKDate & ",presabs from " & strfile1 & "," & rpTables & " where " & _
    strfile1 & ".Empcode = empmst.Empcode and " & strKDate & ">=" & strDTEnc & Format(DateCompDate(typRep.strPeriFr), "dd/mmm/yy") & _
    strDTEnc & " " & strSql & _
    " union select " & strFile2 & ".Empcode," & strFile2 & ".shift,arrtim,deptim,wrkhrs," & strKDate & ",presabs from " & strFile2 & "," & rpTables & " where " & _
    strFile2 & ".Empcode = empmst.Empcode and " & strKDate & "<=" & strDTEnc & Format(DateCompDate(typRep.strPeriTo), "dd/mmm/yy") & strDTEnc & " " & strSql
End If
Select Case bytBackEnd
    Case 1, 2 ''SQLServer,MS-Access
        strGP = strGP & " order by " & strfile1 & ".Empcode," & strfile1 & "." & strKDate & ""
    Case 3    '' ORACLE
        strGP = strGP & " order by Empcode," & strKDate
End Select
dtfromdate = dtFirstDate
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strGP, ConMain, adOpenStatic
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    adrsTemp.MoveFirst
    Do While Not (adrsTemp.EOF) And dtfromdate <= dttodate
        STRECODE = adrsTemp!Empcode
        dtfromdate = dtFirstDate
        p_str = ""
        Do While dtfromdate <= dttodate
            If adrsTemp.EOF Then Exit Do
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtfromdate Then
                    p_str = p_str & IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "") & Spaces(Len(Format(adrsTemp!presabs, "0.00")))
                    If UCase(Trim(adrsTemp!Shift)) = "B" And (adrsTemp!arrtim > 0 Or adrsTemp!deptim > 0) Then
                            SecShft = SecShft + 1
                        End If
                ElseIf adrsTemp!Date <> dtfromdate Then
                    p_str = p_str & Spaces(0)
                End If
            Else
                Exit Do
            End If
             Select Case Left(adrsTemp("presabs"), 2)
                Case pVStar.PrsCode
                    sngpp = sngpp + 0.5
                Case pVStar.AbsCode
                    sngAA = sngAA + 0.5
                Case pVStar.WosCode
                    sngWO = sngWO + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case "OD"
                    sngOD = sngOD + 0.5
                Case "CL"
                    sngCL = sngCL + 0.5
                Case "SL"
                    sngSL = sngSL + 0.5
                Case "ML"
                    sngML = sngML + 0.5
                Case "CO"
                    sngCO = sngCO + 0.5
                Case "PL"
                    sngPL = sngPL + 0.5
                Case Else
                    snglv = snglv + 0.5
            End Select
            Select Case Right(adrsTemp("presabs"), 2)
                Case pVStar.PrsCode
                    sngpp = sngpp + 0.5
                Case pVStar.AbsCode
                    sngAA = sngAA + 0.5
                Case pVStar.WosCode
                    sngWO = sngWO + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case "OD"
                    sngOD = sngOD + 0.5
                Case "CL"
                    sngCL = sngCL + 0.5
                Case "SL"
                    sngSL = sngSL + 0.5
                Case "ML"
                    sngML = sngML + 0.5
                Case "CO"
                    sngCO = sngCO + 0.5
                Case "PL"
                    sngPL = sngPL + 0.5
                Case Else
                    snglv = snglv + 0.5
            End Select
            paid = sngpp + sngWO + sngHL + sngOD + sngCL + sngSL + sngML + sngCO + sngPL
            If dtfromdate = adrsTemp!Date Then adrsTemp.MoveNext
            dtfromdate = DateAdd("d", 1, dtfromdate)
            
        Loop
        
        If Trim(p_str) = "" Then
        Else
            valFld = valFld & IIf(SecShft <> 0, Format(CStr(SecShft), "0.00"), "0.00") & IIf(Len(CStr(SecShft)) >= 6, "", Spaces(Len(Format(CStr(SecShft), "0.00"))))
            ConMain.Execute "insert into " & strRepFile & "" & _
            "(Empcode," & strKDate & ",presabsstr,arrstr,depstr,latestr,earlstr,workstr,otstr,sumlate,sumearly,sumwork,sumot)  values('" & STRECODE & _
            "','" & DTESTR & "','" & p_str & "','" & valFld & "','" & sngpp & "','" & sngAA & "','" & sngWO & "','" & sngHL & "','" & sngPL & "','" & sngOD & "','" & paid & "','" & sngCL & "','" & sngSL & "')"
        End If
        SecShft = 0: valFld = "": dtfromdate = dtFirstDate
        sngpp = 0: sngAA = 0: sngWO = 0: sngHL = 0
        snglv = 0: sngOD = 0: sngCL = 0: sngSL = 0: sngPL = 0: sngCO = 0: sngML = 0
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Periodic Attendance :: Reports")
    Attendance = False
End Function

Public Function Contpredly(presdays As Integer)           '
On Error GoTo ERR_P
Contpredly = True
Dim dtfromdate As Date, dttodate As Date, dtTempDate As Date
Dim PresArray(31) As String
Dim DateArray(31) As Integer
Dim ArrFld As Variant, arrstr As Variant
Dim bytCnt As Byte
Dim STRECODE As String, strTrnFile As String, strTrnFile2 As String, valFld As String, strFld As String
Dim i, j, cnt
Dim Total As Single
Dim FlgS As Boolean

dtfromdate = DateCompDate(frmReports.txtDaily.Text)
dttodate = DateCompDate(frmReports.txtDaily.Text) + (presdays - 1)
strTrnFile = MakeName(MonthName(Month(dtfromdate)), Year(dtfromdate), "trn")
strTrnFile2 = MakeName(MonthName(Month(dttodate)), Year(dttodate), "trn")

If strTrnFile = strTrnFile2 Then
     If adrsTemp.State = 1 Then adrsTemp.Close
     Select Case InVar.strSer
        Case 1, 2 'SQL-Server,MS Access
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and arrtim>0 and " & strKDate & ">=" & strDTEnc & Format(DateCompStr(dtfromdate), "dd/mmm/yy") & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & Format(DateCompStr(dttodate), "dd/mmm/yy") & strDTEnc & "  And presabs IN('P P ','HLHL','WOWO','P A ','A P ','ODOD','ODP','POD') " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
        Case 3 'Oracle
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & "  And presabs IN('P P ','HLHL','WOWO','P A ','A P ') " & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
    End Select
Else
    If adrsTemp.State = 1 Then adrsTemp.Close
    Select Case InVar.strSer
        Case 1, 2 'SQL-server,MS Access
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And presabs IN('P P ','HLHL','WOWO','P A ','A P ') " & strSql & _
        " Union select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile2 & ".shift,OTConf from " & strTrnFile2 & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile2 & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And presabs IN('P P ','HLHL','WOWO','P A ','A P ')" & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
        
     Case 3 ''Oracle
        adrsTemp.Open "select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile & ".shift,OTConf from " & strTrnFile & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And presabs IN('P P ','HLHL','WOWO','P A ','A P ') " & strSql & _
        " Union select empmst.Empcode," & strKDate & ",arrtim,deptim,latehrs,earlhrs,wrkhrs,ovtim," & _
        "presabs," & strTrnFile2 & ".shift,OTConf from " & strTrnFile2 & "," & rpTables & " where " & _
        " empmst.Empcode = " & strTrnFile2 & ".Empcode " & " and " & strKDate & ">=" & strDTEnc & DateCompStr(typRep.strPeriFr) & _
        strDTEnc & " and " & strKDate & "<=" & strDTEnc & DateCompStr(typRep.strPeriTo) & strDTEnc & " And presabs IN('P P ','HLHL','WOWO','P A ','A P ')" & strSql & _
        " order by empmst.Empcode," & strKDate & "", ConMain, adOpenStatic
End Select
End If
arrstr = Array("shf", "Rem")
If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    dtTempDate = dtfromdate
    Do While Not (adrsTemp.EOF) 'And dtTempDate <= dttodate
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtfromdate
        cnt = 0
        Do While dtTempDate <= dttodate And Not (adrsTemp.EOF)
            If adrsTemp!Empcode = STRECODE Then
                If adrsTemp!Date = dtTempDate Then
                    DateArray(cnt) = Day(dtTempDate)
                    PresArray(cnt) = IIf(Not IsNull(adrsTemp!presabs), adrsTemp!presabs, "")
                    cnt = cnt + 1
                Else
                    DateArray(cnt) = 0
                    PresArray(cnt) = ""
                    cnt = cnt + 1
                End If
            Else
                Exit Do
            End If
            If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
            dtTempDate = DateAdd("d", 1, dtTempDate)
       Loop
        valFld = "'" & STRECODE & "'"
        
        Dim Temp As Integer, start As Integer, k As Integer
        strFld = "Empcode"
        bytCnt = 0
            start = 0: Temp = 0
            For i = 0 To DateDiff("d", dtfromdate, dttodate) + 1
                Temp = i
                If DateArray(i) = 0 Then
                    If bytCnt >= presdays Then
                       bytCnt = 0
                       
                    Else
                        For k = start To Temp
                            PresArray(k) = ""
                        Next k
                        bytCnt = 0
                    End If
                    start = Temp + 1
                 Else
                    bytCnt = bytCnt + 1
                End If
            Next i
            
     For j = 0 To UBound(arrstr)
        For i = 1 To DateDiff("d", dtfromdate, dttodate) + 1
            strFld = strFld & "," & arrstr(j) & i
        Next i
     Next j
     
    For i = 0 To DateDiff("d", dtfromdate, dttodate)
        If PresArray(i) <> "" Then
            FlgS = True
            Exit For
        End If
    Next i
    If FlgS = True Then
    dtTempDate = dtfromdate
     For j = 1 To 2
         For i = 0 To DateDiff("d", dtfromdate, dttodate)
             If j = 1 Then
                 valFld = valFld & ",'" & Day(dtTempDate) & "'"
                 dtTempDate = DateAdd("d", 1, dtTempDate)
             ElseIf j = 2 Then
                 valFld = valFld & ",'" & PresArray(i) & "'"
             End If
         Next i
     Next j
     ConMain.Execute " insert into " & strRepMfile & "(" & strFld & ") values(" & valFld & ")"
    End If
    valFld = "": Erase DateArray: Erase PresArray: FlgS = False
    If adrsTemp.EOF Then Exit Do
    'adrsTemp.MoveNext
Loop
End If
Exit Function
ERR_P:
    ShowError ("Continuous present Daily :: Reports")
    Contpredly = False
End Function

Public Function yrAbsPrs1() As Boolean
On Error GoTo ERR_P
yrAbsPrs1 = True
Dim strAPCnt As String, strYrVal As String, strCurYear As String, strCurMon As String
Dim sngTotVal As Single, strAbPrS As String
Dim strFdate As String, strLdate As String, strYearAbPr As String
Dim strArrMon(1 To 12) As String, STRECODE As String, strFileName As String

'
Dim strArrVal(1 To 12) As Single

strArrMon(1) = "January": strArrMon(2) = "February": strArrMon(3) = "March"
strArrMon(4) = "April": strArrMon(5) = "May": strArrMon(6) = "June"
strArrMon(7) = "July": strArrMon(8) = "August": strArrMon(9) = "September"
strArrMon(10) = "October": strArrMon(11) = "November": strArrMon(12) = "December"

If typOptIdx.bytYer = 0 Then
    strAbPrS = pVStar.AbsCode  ' Absent
ElseIf typOptIdx.bytYer = 16 Then
    strAbPrS = "lt_hrs"
ElseIf typOptIdx.bytYer = 17 Then
    strAbPrS = "ot_hrs"
ElseIf typOptIdx.bytYer = 18 Then
    strAbPrS = "LunchLt_Hrs"
Else
    strAbPrS = pVStar.PrsCode  ' Present
End If
strFdate = FdtLdt(CByte(pVStar.Yearstart), pVStar.YearSel, "F")
strLdate = FdtLdt(CByte(pVStar.Yearstart) - 1, IIf(pVStar.YearSel = "1", _
    typRep.strYear, pVStar.YearSel + 1), "L")
strCurMon = strFdate

If Val(pVStar.Yearstart) > Month(DateCompDate(strFdate)) Then
    strFileName = "lvtrn" & Right(CStr(CInt(typRep.strYear) - 1), 2)
Else
    strFileName = "lvtrn" & Right(typRep.strYear, 2)
End If

If adrsEmp.State = 1 Then adrsEmp.Close
adrsEmp.Open "select DISTINCT EMPCODE from " & rpTables & " where joindate <= " & _
strDTEnc & DateCompStr(strLdate) & strDTEnc & " " & strSql & " ORDER BY EMPCODE", ConMain, adOpenStatic
If Not (adrsEmp.EOF And adrsEmp.BOF) Then
    strCurMon = strFdate
    strCurYear = strYearFrom(strFdate)
    YearStr = strYearAbPr 'ASSIGNING VALUE TO GLOBAL VARIABLE -TO USE IN DSR
    Do While Not adrsEmp.EOF
        STRECODE = adrsEmp!Empcode
        strYrVal = ""
        sngTotVal = 0
        Dim i As Integer
         For i = 1 To 12
          strArrVal(i) = 0
         Next i

        strCurMon = strFdate
        strCurYear = strYearFrom(strFdate)
        If Val(pVStar.Yearstart) > Month(DateCompDate(strFdate)) Then
            strFileName = "lvtrn" & Right(CStr(CInt(typRep.strYear) - 1), 2)
        Else
            strFileName = "lvtrn" & Right(typRep.strYear, 2)
        End If

            For i = 1 To 12
            strAPCnt = YearCount1(strFileName, strAbPrS, STRECODE, strCurMon)
            sngTotVal = sngTotVal + Val(strAPCnt)
            If strAPCnt <> "0" Then
                strArrVal(i) = CSng(strAPCnt)
                'strYrVal = strYrVal & Spaces(Len(strAPCnt)) & strAPCnt
            Else
                strYrVal = 0
            End If
            strCurMon = CStr(DateAdd("m", 1, DateCompDate(strCurMon)))
            If Month(DateCompDate(strCurMon)) <= Month(DateCompDate(strLdate)) _
            And pVStar.Yearstart <> 1 Then
                strCurYear = Year(DateCompDate(strFdate)) + 1
            End If
        Next i
        strYrVal = strYrVal & Spaces(3) & IIf(sngTotVal <= 0, "", Spaces(Len(Format(sngTotVal, "0.00"))) & Format(sngTotVal, "0.00"))
        If Trim(strYrVal) <> "" Then
            If typOptIdx.bytYer = 0 And GetFlagStatus("FANANCIALYEAR") Then
            ConMain.Execute "insert into " & strRepFile & " values ( '" & STRECODE & "'," & strArrVal(10) & "," & strArrVal(11) & "," & strArrVal(12) & "," & strArrVal(1) & "," & strArrVal(2) & "," & strArrVal(3) & "," & strArrVal(4) & "," & strArrVal(5) & "," & strArrVal(6) & "," & strArrVal(7) & "," & strArrVal(8) & "," & strArrVal(9) & "," & sngTotVal & ")"
            Else
            ConMain.Execute "insert into " & strRepFile & " values ( '" & STRECODE & "'," & strArrVal(1) & "," & strArrVal(2) & "," & strArrVal(3) & "," & strArrVal(4) & "," & strArrVal(5) & "," & strArrVal(6) & "," & strArrVal(7) & "," & strArrVal(8) & "," & strArrVal(9) & "," & strArrVal(10) & "," & strArrVal(11) & "," & strArrVal(12) & "," & sngTotVal & ")"
            End If
'            conmain.Execute "insert into " & strRepFile & " values ( '" & STRECODE & "',0,0,0," & strArrVal(4) & "," & strArrVal(5) & "," & strArrVal(6) & "," & strArrVal(7) & "," & strArrVal(8) & "," & strArrVal(9) & "," & strArrVal(10) & "," & strArrVal(11) & "," & strArrVal(12) & "," & sngTotVal & ")"
        End If
        adrsEmp.MoveNext
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Yearly Absent Present :: Reports")
    yrAbsPrs1 = False
End Function

Public Function Monthsummery()                 ' for Nestale Samalkha
    On Error GoTo ERR_P
    Monthsummery = True
    Dim rsEmp As New ADODB.Recordset
    Dim strTrnFile As String
    Dim absent As Double
    strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "trn")
    If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT deptdesc." & strKDesc & ", COUNT(*) AS TotEmp fROM Empmst INNER JOIN deptdesc ON Empmst.dept = deptdesc.dept GROUP BY deptdesc." & strKDesc & " ORDER BY deptdesc." & strKDesc & "", ConMain, adOpenStatic
        Do While Not (adrsTemp.EOF)
            ConMain.Execute " insert into " & strRepMfile & "(Rem1,Rem2) values ('" & adrsTemp.Fields(0) & "'," & adrsTemp.Fields(1) & ")", i
            adrsTemp.MoveNext
        Loop
                        
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'P ' And Right(" & strTrnFile & ".presabs, 2) = 'P ') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                ConMain.Execute "update " & strRepMfile & " set Rem3=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'A ' And Right(" & strTrnFile & ".presabs, 2) = 'A ') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
             ConMain.Execute "update " & strRepMfile & " set Rem4=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'WO' And Right(" & strTrnFile & ".presabs, 2) = 'WO') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic  'run for oracle only
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                ConMain.Execute "update " & strRepMfile & " set Rem5=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'CL' And Right(" & strTrnFile & ".presabs, 2) = 'CL') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                 ConMain.Execute "update " & strRepMfile & " set Rem6=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'SL' And Right(" & strTrnFile & ".presabs, 2) = 'SL') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                 ConMain.Execute "update " & strRepMfile & " set Rem7=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'CO' And Right(" & strTrnFile & ".presabs, 2) = 'CO') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                 ConMain.Execute "update " & strRepMfile & " set Rem8=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'OD' And Right(" & strTrnFile & ".presabs, 2) = 'OD') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                 ConMain.Execute "update " & strRepMfile & " set Rem9=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'EL' And Right(" & strTrnFile & ".presabs, 2) = 'EL') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                ConMain.Execute "update " & strRepMfile & " set Rem10=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'HL' And Right(" & strTrnFile & ".presabs, 2) = 'HL') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                ConMain.Execute "update " & strRepMfile & " set Rem11=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'SP' And Right(" & strTrnFile & ".presabs, 2) = 'SP') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                 ConMain.Execute "update " & strRepMfile & " set Rem12=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'LP' And Right(" & strTrnFile & ".presabs, 2) = 'LP') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                 ConMain.Execute "update " & strRepMfile & " set Rem13=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
        
Exit Function
ERR_P:
    ShowError ("monJublientReport :: Reports")
    Monthsummery = False
End Function
Public Function shiftroll() As Boolean   '' for Suzlon
On Error GoTo ERR_P
shiftroll = True
Dim strTrnFile As String, strlvfile As String, p_str As String, strFld As String, valFld As String
Dim STRECODE As String, strqry As String, strLv As String, arrLv() As String, strCat As String
Dim adrsLvCD As New ADODB.Recordset
Dim dtFDate As Date, dtLDate As Date, dtTempDate As Date
Dim LvP As Single, LvUP As Single, Total As Single, totab As Single
Dim i As Integer

strTrnFile = MakeName(typRep.strMonMth, typRep.strMonYear, "Shf")
If Val(pVStar.Yearstart) > MonthNumber(typRep.strMonMth) Then
    strlvfile = "lvtrn" & Right(CStr(CInt(typRep.strMonYear) - 1), 2)
Else
    strlvfile = "lvtrn" & Right(typRep.strMonYear, 2)
End If
If Not FindTable(strlvfile) Then
    shiftroll = False
    MsgBox NewCaptionTxt("M7005", adrsMod), vbInformation
    Exit Function
End If
dtFDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "F"))
dtLDate = DateCompDate(FdtLdt(MonthNumber(typRep.strMonMth), typRep.strMonYear, "l"))
dtTempDate = dtFDate

strqry = "select empmst.empcode," & strTrnFile & ".*,empmst.cat from " & strTrnFile & "," & _
         rpTables & " where " & strTrnFile & ".empcode = empmst.empcode " & strSql & _
         " order by " & strTrnFile & ".empcode"

If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open strqry, ConMain, adOpenStatic

If Not (adrsTemp.BOF And adrsTemp.EOF) Then
    DateStr = monDateStr(Day(dtLDate))
    Do While Not (adrsTemp.EOF) 'And dtTempDate < dtLDate
        frmReports.Refresh
        STRECODE = adrsTemp!Empcode
        dtTempDate = dtFDate
        p_str = ""
        totab = 0
        strFld = "Empcode"
        Do While dtTempDate <= dtLDate And Not (adrsTemp.EOF)
                If adrsTemp!Empcode = STRECODE Then
                strFld = strFld & ",Rem" & Day(dtTempDate) '
                If adrsTemp!Date = dtTempDate Then
                   p_str = p_str & "'" & IIf(Not IsNull(adrsTemp!Shift), adrsTemp!Shift, _
                    "") & Spaces(0) & "',"
                ElseIf adrsTemp!Date <> dtTempDate Then
                   p_str = p_str & "'" & Spaces(0) & "',"
                End If
            Else
                Exit Do
            End If
           If dtTempDate = adrsTemp!Date Then adrsTemp.MoveNext
        dtTempDate = DateAdd("d", 1, dtTempDate)

        Loop
        p_str = Mid(p_str, 1, Len(p_str) - 1)
        valFld = "'" & STRECODE & "'"
        If Trim(p_str) = "" Then
        Else
           ConMain.Execute "insert into " & strRepMfile & _
           "(" & strFld & ",shf1,shf2,shf3,shf4) values(" & valFld & "," & p_str & ",'" & Format(CStr(LvP), "0.00") & "','" & Format(CStr(LvUP), "0.00") & "','" & Format(CStr(Total), "0.00") & "','" & Format(CStr(totab), "0.00") & "')"
         p_str = ""
       End If
    Loop
End If
Exit Function
ERR_P:
    ShowError ("Form XI:: Reports")
    shiftroll = False
    'Resume Next
End Function

Public Function shiftabsent() As Boolean
On Error GoTo RepErr
shiftabsent = True
Dim rsEmp As New ADODB.Recordset
    Dim strTrnFile As String
    Dim absent As Double
    strTrnFile = Left(MonthName(Month(CDate(typRep.strDlyDate))), 3) & Right(Year(CDate(typRep.strDlyDate)), 2) & "trn"
    If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT deptdesc." & strKDesc & ", COUNT(*) AS TotEmp fROM Empmst INNER JOIN deptdesc ON Empmst.dept = deptdesc.dept GROUP BY deptdesc." & strKDesc & " ORDER BY deptdesc." & strKDesc & "", ConMain, adOpenStatic
        Do While Not (adrsTemp.EOF)
            ConMain.Execute " insert into " & strRepFile & "(Empcode,Present) values ('" & adrsTemp.Fields(0) & "','" & adrsTemp.Fields(1) & "')", i
            adrsTemp.MoveNext
        Loop
                        
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "SELECT empmst.dept, deptdesc." & strKDesc & ", count(" & strTrnFile & ".presabs) AS cnt, " & strTrnFile & ".presabs From empmst, deptdesc, " & strTrnFile & " Where empmst.dept = deptdesc.dept And empmst.Empcode = " & strTrnFile & ".Empcode And (Left(" & strTrnFile & ".presabs, 2) = 'P ' And Right(" & strTrnFile & ".presabs, 2) = 'P ') GROUP BY empmst.dept, deptdesc.[desc]," & strTrnFile & ".presabs", ConMain, adOpenStatic, adLockOptimistic
        
        If Not (adrsTemp.BOF And adrsTemp.EOF) Then
            Do While Not (adrsTemp.EOF)
                ConMain.Execute "update " & strRepFile & " set Rem3=" & adrsTemp.Fields("cnt") & " where Rem1='" & adrsTemp.Fields("desc") & "'"
                adrsTemp.MoveNext
            Loop
        End If
    
Exit Function
RepErr:
        ShowError ("shiftabsent:: Reports")
        shiftabsent = False
        'Resume Next
End Function


