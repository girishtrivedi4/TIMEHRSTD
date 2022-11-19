Attribute VB_Name = "mdlLog"
'' Module for LOG
'' --------------
'' Steps to be taken to Adjust Log Concept in the Working Copy.
'' 01. Add a Data Environment and Name it as DELOG
'' 02. Name the Connection to cnLog.
'' 03. Call the ConnectLog Boolean Function in the MAIN immidiately after Connection
'' 04. If the Connection is Succesfull Set the Recordsets and get the highest Record Number.
'' 05. On Application Terminate Save the Highest Log Number in MDI_Unload Event of the MDI.
Option Explicit
'' Log Recordsets
Public adrsLog As New ADODB.Recordset
Public Con As New ADODB.Connection

Public Function ConnectLog() As Boolean     '' Function to Connect to the Log MDB
On Error GoTo ERR_P
'Dim TMP As New ADODB.Connection
'TMP.ConnectionString =        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\IVAttendo\Data\AttendoLog.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ATTENDOLOG"
'TMP.Open

Call WriteLog("Step1 connectlog")
If ConLog.State = 1 Then ConLog.Close
Call WriteLog("Step2 Data Source")
ConLog.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Data\TimeHRLog.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ATTENDOLOG"
'conlog.Properties("Data Source") = App.Path & "\Data\AttendoLog.Mdb"
''Call WriteLog("Step3 EX Pro")
'conlog.Properties("Extended Properties") = "Microsoft.Jet.OLEDB.4.0:Database Password = ATTENDOLOG"
'Call WriteLog("Step4 Open")
ConLog.Open
Call SetLogRecs
ConnectLog = True
Exit Function
ERR_P:
    Call WriteLog(Err.Description)
    ConnectLog = False
End Function

Public Sub GetLogRecordNumber()             '' Gets the Highest RecordNumber
On Error GoTo ERR_P
If adrsLog.State = 1 Then adrsLog.Close
adrsLog.Open "Select RecNumber from RECNUM"
If Not (adrsLog.EOF And adrsLog.BOF) Then
    If IsNull(adrsLog("RecNumber")) Then
        typLog.lngRecord = 0
    Else
        typLog.lngRecord = adrsLog("RecNumber")
    End If
Else
    typLog.lngRecord = 0
End If
Exit Sub
ERR_P:
    ShowError ("Error Getting Log Details :: " & vbCrLf & _
    "It is Recommended that you do not Proceed with the Application :: Log")
    typLog.lngRecord = 0
End Sub

Private Sub SetLogRecs()        '' Sets the Recordsets Created for LOG
On Error Resume Next
adrsLog.ActiveConnection = ConLog
adrsLog.CursorType = adOpenStatic
End Sub

Public Sub AddActivityLog(ByVal strModeType As String, _
bytTranType As Byte, bytTranSource As Byte)         '' Procedure to Add the Activity Log
On Error GoTo ERR_P
typLog.strModeType = strModeType                    '' Mode Type
typLog.bytTranType = bytTranType                    '' Transaction Type
typLog.bytTranSource = bytTranSource                '' Transaction Source
typLog.lngRecord = typLog.lngRecord + 1             '' Increment the Log Record Number
typLog.strDate = DateSaveIns(CStr(Date))            '' Date
typLog.sngTime = IIf(Hour(Time) = 0, 24, Hour(Time)) & "." & Format(Minute(Time), "00") '' Time
typLog.strUsername = IIf(UCase(UserName) = UCase(strPrintUser), "*****", UserName) '' User Name
ConLog.Execute "insert into Activity Values(" & typLog.sngTime & ",'" & _
typLog.strUsername & "'," & typLog.lngRecord & ",'" & typLog.strModeType & "'," & _
typLog.bytTranType & "," & typLog.bytTranSource & ",#" & typLog.strDate & "#)"
Exit Sub
ERR_P:
    ShowError ("Failed to log the Activity :: Log")
End Sub
Public Sub AuditInfo(ByVal CmdStr As String, ByVal FrmStr As String, ByVal LogMsg As String, Optional tempuser As String, Optional InTime As Single, Optional outTime As Single, Optional OldIn As Single, Optional OldOut As Single, Optional dte As String)     ' 29-05
On Error GoTo ERR_P
If GetFlagStatus("NOMIS2010") = False Or GetFlagStatus("ANSA") = True Then
    If UserName = "" Then
        typAudit.strUser = tempuser
    Else
        typAudit.strUser = IIf(UCase(UserName) = UCase(strPrintUser), "*****", UserName)  '' User Name
    End If
    typAudit.strDt = DateSaveIns(CStr(Date))            '' Date
    typAudit.sngTmt = IIf(Hour(Time) = 0, "00", Hour(Time)) & "." & Format(Minute(Time), "00") '' Time
    typAudit.strMsg = LogMsg
    typAudit.lngRecNum = typAudit.lngRecNum + 1
    typAudit.strIn = Format(InTime, "00.00")
    typAudit.strout = Format(outTime, "00.00")
'    typAudit.strIp = frmAuditInfo.winip.LocalIP
    typAudit.dte = dte

End If
Exit Sub
ERR_P:
    ShowError ("Failed to log Audit")
    Resume Next
End Sub

Public Function NewCaptionTxt(strCapId As String, adrsTmp As ADODB.Recordset, Optional lngTmpSpace As Long = 1)
'On Error Resume Next
adrsTmp.MoveFirst
adrsTmp.Find "Id='" & strCapId & "'"
NewCaptionTxt = adrsTmp(strCapField) & Space(lngTmpSpace)
If NewCapFlag Then  ' 06-08
    NewCaptionTxt = RepNewCaptionTxt(NewCaptionTxt)
End If
End Function
Public Function RepNewCaptionTxt(ByVal strTemp As String)   'Added by  06-08
On Error GoTo Err
 Dim i As Integer
 For i = 1 To UBound(LookFor)
    RepNewCaptionTxt = Replace(strTemp, LookFor(i), RepWith(i))
    strTemp = RepNewCaptionTxt
 Next
 Exit Function
Err:
 ShowError ("RepNewCaptionTxt:" & Err.Description)
 'Resume Next
End Function

Public Sub LoginLog() '
If Hardboot <> "T" Then
Dim rs As New ADODB.Recordset
    If Not LoginStatus Then Msr_no = 0
rs.Open "Select Max(ErrorId) from ErrorLog ", ConMain, adOpenKeyset, adLockReadOnly
If IsNull(rs.Fields(0)) Then
    Msr_no = 1
Else
    If Not LoginStatus Then
        Msr_no = rs.Fields(0) + 1
    End If
End If
If Not LoginStatus Then
    ConLog.Execute "Update TransDet set TransType = " & Msr_no & ", TransDesc='T' Where TransSource=35"
    ConMain.Execute "insert into errorLog (ErrorName,ErrorId) values ('" & UserName & "', " & Msr_no & ")"
End If
End If
End Sub

Public Sub BootStatus() '
Dim rs1 As New ADODB.Recordset
rs1.Open " Select TransDesc,TransType from TransDet where TransSource=35", ConLog
If IsNull(rs1.Fields(0)) = True Then
    Hardboot = ""
Else
    Hardboot = rs1.Fields(0)
End If
If Hardboot = "T" Then
    If InVar.strSer = 3 Then
        'this is previous code
        'conmain.Execute "Delete * from errorlog where errorid=" & rs1.Fields(1)
        'this is new code add by  MIS2007DF019
        ConMain.Execute "Delete from errorlog where errorid=" & rs1.Fields(1)
    Else
        ConMain.Execute "Delete from errorlog where errorid=" & rs1.Fields(1)
    End If
End If
End Sub

