VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImpTL 
   Caption         =   "Import"
   ClientHeight    =   2475
   ClientLeft      =   4200
   ClientTop       =   2985
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2160
      ScrollBars      =   1  'Horizontal
      TabIndex        =   9
      Text            =   " "
      Top             =   120
      Width           =   4035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6735
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   1290
         TabIndex        =   4
         Tag             =   "D"
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   3450
         TabIndex        =   3
         Tag             =   "D"
         Text            =   "  "
         Top             =   750
         Width           =   1215
      End
      Begin VB.TextBox txtECode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Tag             =   "D"
         Text            =   " "
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtEName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   3480
         TabIndex        =   1
         Tag             =   "D"
         Text            =   " "
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   2880
         TabIndex        =   8
         Top             =   855
         Width           =   195
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   390
      End
      Begin VB.Label lblECode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Code"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   300
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImpData 
      Caption         =   "ImportData"
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "Import"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdImpL 
      Caption         =   "ImportLeave"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblPath 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Path Of Excel File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmImpTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset
Dim strCatAvail As String, strF As String, strT As String, LvOpt As String
Dim dtJoin As Date, Fdate As Date, TDate As Date
Dim TotDays As Single, sngDaysBal As Single
Dim blnUnPaid As Boolean, blnNoBal As Boolean
Dim bytMin As Byte, bytMax As Byte
Dim intTimes As Integer
Dim ECode As String, LCode As String, EName As String, strHf_Opt As String, strRW As String
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdImp_Click()
    Screen.MousePointer = vbHourglass
    If Not ImpExcel Then Exit Sub

        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "Select  distinct lcode from ImportLeave", VstarDataEnv.cnDJConn
        Do While Not adrsTemp.EOF
            If adrsDept1.State = 1 Then adrsDept1.Close
            adrsDept1.Open "Select  distinct * from leavdesc where lvcode='" & adrsTemp.Fields("lcode") & "'", VstarDataEnv.cnDJConn
            If adrsDept1.EOF And adrsDept1.BOF Then
                MsgBox adrsTemp.Fields("lcode") & " Leave is not present in Leave Master."
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
            adrsTemp.MoveNext
        Loop

    If Not ProcessDT Then Exit Sub
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdImpData_Click()
CD1.Flags = cdlOFNHideReadOnly
CD1.Filter = "Excel Files|*.xls|All Files|*.*"
CD1.FilterIndex = 1
CD1.FileName = ""
CD1.ShowOpen
Call ImportData(CD1.FileName)
End Sub

Private Sub cmdPath_Click()
    CD1.Flags = cdlOFNHideReadOnly
    CD1.Filter = "Excel Files|*.xls"
    CD1.FilterIndex = 1
    CD1.FileName = ""
    CD1.ShowOpen
    txtPath.Text = CD1.FileName
    If CD1.FileName <> "" Then
        cmdImp.Enabled = True
    Else
        cmdImp.Enabled = False
    End If
End Sub

Private Function ImpExcel() As Boolean
On Error GoTo ERR_P
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim rsTemp As New ADODB.Recordset
    Dim rsAccess As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim ECodeSize As Integer
    Dim cn As New ADODB.Connection
    Set cn = New ADODB.Connection
    ImpExcel = True
    Set xlApp = New Excel.Application
    xlApp.Workbooks.Open CD1.FileName
    xlApp.Workbooks(1).Activate
    Set xlSheet = xlApp.ActiveWorkbook.Sheets(1)
     
    If cn.State = 1 Then cn.Close
'    cn.ConnectionString = "Provider=MSDAORA ;Data Source=" & CD1.FileName & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
'    cn.Open
    cn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0 ;Data Source=" & CD1.FileName & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
    cn.Open
    If FindTable("ImportLeave") Then VstarDataEnv.cnDJConn.Execute "drop table ImportLeave"
    Sleep (1000)
    
    Sleep (2000)
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open "select e_codesize from install", VstarDataEnv.cnDJConn, adOpenStatic
    ECodeSize = rsTemp.Fields(0)
    
'        strSql = "insert into ImportLeave select * FROM [Excel 8.0;DATABASE=" & CD1.FileName & ";HDR=Yes].[sheet1$]"
'        VstarDataEnv.cnDJConn.Execute strSql
        strSql = "select * FROM [Excel 8.0;DATABASE=" & CD1.FileName & ";HDR=Yes].[sheet1$]"
        'VstarDataEnv.cnDJConn.Execute strSql
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open strSql, cn, adOpenStatic
        'rsTemp.Open strSql, VstarDataEnv.cnDJConn, adOpenStatic
        
        If rsAccess.State = 1 Then rsAccess.Close
        rsAccess.Open "select * from ImportLeave", VstarDataEnv.cnDJConn, adOpenDynamic, adLockOptimistic
        For i = 1 To rsTemp.RecordCount 'looping through all records of the excel file
            ECode = Trim(xlSheet.Cells(i + 1, 1).Value)
            If Not (ECode = "") Then
                If Len(ECode) <> ECodeSize Then ECode = ZeroPadding(ECodeSize - Len(ECode), ECode)
                Fdate = xlSheet.Cells(i + 1, 2).Value
                TDate = xlSheet.Cells(i + 1, 3).Value
                LCode = xlSheet.Cells(i + 1, 4).Value
                TotDays = xlSheet.Cells(i + 1, 5).Value
                LvOpt = xlSheet.Cells(i + 1, 6).Value
                rsAccess.AddNew
                rsAccess.Fields("Empcode") = ECode
                rsAccess.Fields("FromDate") = Fdate
                rsAccess.Fields("Todate") = TDate
                rsAccess.Fields("LCode") = LCode
                rsAccess.Fields("Days") = TotDays
                rsAccess.Fields("hf_option") = UCase(Trim(LvOpt))
                rsAccess.Update
            End If
            ECode = "": LvOpt = "": TotDays = 0: LCode = ""
            rsTemp.MoveNext
        Next i
    xlApp.DisplayAlerts = False
    xlApp.ActiveWorkbook.Close
    'xlApp.Workbooks.Close
    xlApp.Application.Quit
    Set xlApp = Nothing: Set xlBook = Nothing: Set xlSheet = Nothing
    Exit Function
ERR_P:
    ImpExcel = False: Screen.MousePointer = vbNormal
    ShowError ("ImpExcel:: " & Me.Caption & Err.Description & ":" & Err.Number & ":" & Erl)
    'Resume Next
End Function
Private Function ProcessDT() As Boolean
On Error GoTo ERR_P
    Dim cn As New ADODB.Connection
    Set cn = New ADODB.Connection
    ProcessDT = True
    If adrsEmp.State = 1 Then adrsEmp.Close
        adrsEmp.Open "Select ImportLeave.*,name from ImportLeave,empmst where empmst.empcode=ImportLeave.empcode", VstarDataEnv.cnDJConn

    If (adrsEmp.EOF And adrsEmp.BOF) Then
        MsgBox "No Data Found", vbExclamation
        Exit Function
    Else
    Do While Not adrsEmp.EOF
        strHf_Opt = ""
        If IsNull(adrsEmp.Fields("empcode")) Or adrsEmp.Fields("empcode") = "" Then GoTo NextEmp
        ECode = adrsEmp.Fields("empcode"): EName = adrsEmp.Fields("name")
        Fdate = adrsEmp.Fields("fromdate"): TDate = adrsEmp.Fields("todate")
        TotDays = adrsEmp.Fields("days")
        txtFrom.Text = adrsEmp.Fields("fromdate"): txtFrom.Refresh
        txtTo.Text = adrsEmp.Fields("todate"): txtTo.Refresh
        txtECode.Text = adrsEmp.Fields("empcode"): txtECode.Refresh
        txtEName.Text = adrsEmp.Fields("name"): txtEName.Refresh
        If Not GetCat Then GoTo NextEmp
        If Not ValidDate(txtFrom) Then GoTo NextEmp
        If Not ValidDate(txtTo) Then GoTo NextEmp
        If Not LeaveType() Then GoTo NextEmp
            
            strF = "": strT = ""
            If Len(strHf_Opt) = 4 Then
                strF = Left(strHf_Opt, 2): strT = Right(strHf_Opt, 2)
            ElseIf Len(strHf_Opt) = 2 Then
                strF = Left(strHf_Opt, 1) & " ": strT = Right(strHf_Opt, 1) & " "
            ElseIf Len(strHf_Opt) = 3 And Mid(strHf_Opt, 3, 1) = "T" Then
                strF = Left(strHf_Opt, 2): strT = Right(strHf_Opt, 1) & " "
            ElseIf Len(strHf_Opt) = 3 And Mid(strHf_Opt, 2, 1) = "T" Then
                strF = Left(strHf_Opt, 1) & " ": strT = Right(strHf_Opt, 2)
            End If
            strHf_Opt = strF & strT

        If Not ValidateAddmaster Then GoTo NextEmp
        If Not SaveAddMaster Then GoTo NextEmp
NextEmp:
        adrsEmp.MoveNext
    Loop
            MsgBox ("Imported Sucessfully")
    End If

    Exit Function
ERR_P:
    ProcessDT = False: Screen.MousePointer = vbNormal
    'Resume Next
    ShowError ("ProcessDT:: " & Me.Caption & Err.Description & ":" & Err.Number & ":" & Erl)
End Function

Private Function GetCat() As Boolean
On Error GoTo ERR_P
GetCat = True
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open "Select Cat,JoinDate,COCode from EmpMst where EmpCode='" & ECode & "'", VstarDataEnv.cnDJConn
If (adrsPaid.EOF And adrsPaid.BOF) Then
    MsgBox NewCaptionTxt("07016", adrsC), vbExclamation
    GetCat = False
    Exit Function
Else
    strCatAvail = adrsPaid("Cat")
    If Not IsNull(adrsPaid("JoinDate")) Then
        dtJoin = DateCompDate(adrsPaid("JoinDate"))
    Else
        dtJoin = DateCompDate("31-December-2100")
    End If
    'bytCOCode = IIf(IsNull(adrsPaid("COCode")), 100, adrsPaid("COCode"))
End If
Exit Function
ERR_P:
    GetCat = False
    ShowError ("Getcat :: " & Me.Caption)
End Function

Public Function LeaveType() As Boolean
On Error GoTo ERR_P
LeaveType = True
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select run_wrk from leavdesc where cat='" & strCatAvail & "' and lvcode='" & LCode & "'", VstarDataEnv.cnDJConn
If Not (adrsDept1.EOF And adrsDept1.BOF) Then
    If adrsDept1(0) = "O" Then
        strRW = "R"
    Else
        strRW = adrsDept1(0)
    End If
Else
    LeaveType = False
End If
Exit Function
ERR_P:
    LeaveType = False
    ShowError ("LeaveType")
End Function

Private Sub Form_Load()
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions Where ID Like '07%'", VstarDataEnv.cnDJConn, adOpenStatic
End Sub
Private Function ValidateAddmaster() As Boolean     '' Validate Details befor Availing the
On Error GoTo ERR_P                                 '' Leave
ValidateAddmaster = True
'' Check for Invalid Number of AvailLeave Days
If Val(TotDays) <= 0 Then
    MsgBox NewCaptionTxt("07018", adrsC), vbExclamation
    ValidateAddmaster = False
    Exit Function
End If
'' Check for Invalid date
If Not ValidLeaveDate Then
    ValidateAddmaster = False
    Exit Function
End If
Call GetLeaveDetails
'' Check if Leaves are Already Availed for the Specified Dates
If Not ALreadyAvailedDate Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check if he has Availed Leaves for More than Allowed Times
If Not NumOfTimesAvailed Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check Minimum & Maximum Limits
If Not CheckMinMaxLeave Then
    ValidateAddmaster = False
    Exit Function
End If
'' Check if the Person is Immidiate Absent before or or After the Leave Dates
If Not ImmediateAbsent Then
    ValidateAddmaster = False
    Exit Function
End If
If Not ImmediateLeave Then
    ValidateAddmaster = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ValidateAddMaster :: " & Me.Caption)
    Resume Next
    ValidateAddmaster = False
End Function

Private Function ValidLeaveDate() As Boolean        '' Validates the Leave Dates Specified
ValidLeaveDate = True
'' Check for EmptyDate
If Trim(txtFrom.Text) = "" Then
    MsgBox NewCaptionTxt("00016", adrsMod), vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If txtTo.Text = "" Then
    MsgBox NewCaptionTxt("00017", adrsMod), vbExclamation
    txtTo.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateCompDate(txtFrom.Text) > DateCompDate(txtTo.Text) Then
    MsgBox NewCaptionTxt("00018", adrsMod), vbExclamation
'    txtTo.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) > 11 Or _
DateDiff("m", Year_Start, DateCompDate(txtFrom.Text)) < 0 Then
    MsgBox NewCaptionTxt("00019", adrsMod) & txtFrom.Text & NewCaptionTxt("00021", adrsMod), _
    vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
If DateDiff("m", Year_Start, DateCompDate(txtTo.Text)) > 11 Or _
DateDiff("m", Year_Start, DateCompDate(txtTo.Text)) < 0 Then
    MsgBox NewCaptionTxt("00020", adrsMod) & txtTo.Text & NewCaptionTxt("00021", adrsMod), _
    vbExclamation
    txtTo.SetFocus
    ValidLeaveDate = False
End If
    Exit Function

If DateCompDate(txtFrom.Text) < dtJoin Then
    MsgBox NewCaptionTxt("00112", adrsMod), vbExclamation
    txtFrom.SetFocus
    ValidLeaveDate = False
    Exit Function
End If
End Function

Private Sub GetLeaveDetails()       '' Gets the Other Primary Details of a Leave
On Error GoTo ERR_P                 '' from the Leave Master
If adrsPaid.State = 1 Then adrsPaid.Close
    adrsPaid.Open "Select AllowDays,MinAllowDays,No_OfTimes,Paid,Type from leavdesc where LvCode='" & LCode & "' and Cat='" & strCatAvail & "'", VstarDataEnv.cnDJConn
bytMax = IIf(IsNull(adrsPaid("AllowDays")), 0, adrsPaid("AllowDays"))
bytMin = IIf(IsNull(adrsPaid("MinAllowDays")), 0, adrsPaid("MinAllowDays"))
intTimes = IIf(IsNull(adrsPaid("No_OfTimes")), 0, adrsPaid("No_OfTimes"))
blnUnPaid = IIf(adrsPaid("Paid") = "N", True, False)
blnNoBal = IIf(adrsPaid("Type") = "N", True, False)
Exit Sub
ERR_P:
    ShowError ("GetLeaveDetails :: " & Me.Caption)
    bytMax = 0
    bytMin = 0
    intTimes = 0
End Sub
Private Function ALreadyAvailedDate() As Boolean    '' Checks if Leave is Already Availed
On Error GoTo ERR_P                                 '' by the Employee for the Same Dates
ALreadyAvailedDate = True
Dim strA_R As String, strHFOPT As String
Dim bytCtr As Byte
strA_R = "select * from LvInfo" & Right(pVStar.YearSel, 2) & " where ((" & strDTEnc & _
Format(DateCompDate(txtFrom.Text), "dd/MMM/yy") & strDTEnc & " between fromdate and todate ) or (" & _
strDTEnc & Format(DateCompDate(txtTo.Text), "dd/MMM/yy") & strDTEnc & _
" between fromdate and todate)) and trcd=4 and Lcode='" & LCode & "' and Empcode='" & ECode & "'"
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open strA_R, VstarDataEnv.cnDJConn
bytCtr = 0
If Not (adrsPaid.EOF And adrsPaid.BOF) Then
        Do While Not adrsPaid.EOF
                If strF = "F " Then bytCtr = bytCtr + 1
                If strT = "T " Then bytCtr = bytCtr + 1
                strHFOPT = Left(adrsPaid("hf_option"), 2)
                If strHFOPT = strF Or strHFOPT = "F " Then bytCtr = bytCtr + 1
                strHFOPT = Right(adrsPaid("HF_Option"), 2)
                If strHFOPT = strT Or strHFOPT = "T " Then bytCtr = bytCtr + 1
                adrsPaid.MoveNext
        Loop
End If
If bytCtr > 0 Then
        MsgBox NewCaptionTxt("07027", adrsC), vbExclamation
        ALreadyAvailedDate = False
        Exit Function
End If
Exit Function
ERR_P:
    ShowError ("ALreadyAvailedDate :: " & Me.Caption)
    ALreadyAvailedDate = False
End Function
Private Function NumOfTimesAvailed() As Boolean     '' Checks if How Many times the Employee
On Error GoTo ERR_P                                 '' is Allowed to Avail the Leave
NumOfTimesAvailed = True
Dim bytCntTmp As Byte, bytTmp As Byte
bytTmp = 0
'' Number of Times Availed
If adrsDept1.State = 1 Then adrsDept1.Close
adrsDept1.Open "Select * from lvinfo" & Right(pVStar.YearSel, 2) & " where LCode='" & LCode & "' And trcd=4 and Empcode='" & ECode & "'" & _
    " Order by  LCode,Fromdate", VstarDataEnv.cnDJConn, adOpenStatic
bytTmp = adrsDept1.RecordCount
If intTimes > 0 Then
    If bytTmp >= intTimes Then
        If MsgBox(NewCaptionTxt("07019", adrsC) & intTimes & NewCaptionTxt("07020", adrsC), _
        vbQuestion + vbYesNo) = vbYes Then
            NumOfTimesAvailed = True
        Else
            txtFrom.SetFocus
            NumOfTimesAvailed = False
        End If
    End If
End If
Exit Function
ERR_P:
    ShowError ("NumOfTimesAvailed :: " & Me.Caption)
    NumOfTimesAvailed = False
End Function
Private Function CheckMinMaxLeave() As Boolean  '' Checks the Minimum & the Maximum
CheckMinMaxLeave = True                         '' Leaves the Employee is Allowed to Avail
If bytMax <> 0 Then
    If Val(TotDays) > bytMax Then
        If MsgBox(NewCaptionTxt("07021", adrsC) & bytMax & NewCaptionTxt("07022", adrsC), _
        vbQuestion + vbYesNo) = vbYes Then
            CheckMinMaxLeave = True
            Exit Function
        Else
            txtFrom.SetFocus
            CheckMinMaxLeave = False
            Exit Function
        End If
    End If
End If
If bytMin <> 0 Then
    If Val(TotDays) < bytMin Then
        If MsgBox(NewCaptionTxt("07023", adrsC) & bytMin & NewCaptionTxt("07022", adrsC), _
        vbQuestion + vbYesNo) = vbYes Then
            CheckMinMaxLeave = True
            Exit Function
        Else
            txtFrom.SetFocus
            CheckMinMaxLeave = False
            Exit Function
        End If
    End If
End If
End Function
Private Function ImmediateAbsent() As Boolean       '' Checks if the Employee is Absent on
On Error GoTo ERR_P                                 '' Consecutive Days
Dim strTmp As String, strPATmp As String, bytPATmp As Byte
ImmediateAbsent = True
strTmp = ""
strPATmp = ""
bytPATmp = 0
strTmp = GetMnlTrnFile((DateCompDate(txtFrom.Text) - 1))
If FindTable(strTmp) Then
    If adrsPaid.State = 1 Then adrsPaid.Close
    adrsPaid.Open "Select Presabs from " & strTmp & " where " & strKDate & " =" & strDTEnc & _
    DateCompStr(CStr(DateCompDate(txtFrom.Text) - 1)) & strDTEnc & " and EmpCode=" & _
    "'" & ECode & "'", VstarDataEnv.cnDJConn, adOpenKeyset
    If Not (adrsPaid.BOF And adrsPaid.EOF) Then
        strPATmp = adrsPaid("Presabs")
        bytPATmp = Len(strPATmp)
        If Mid(strPATmp, 1, bytPATmp / 2) = pVStar.AbsCode Then
            Select Case MsgBox(NewCaptionTxt("07028", adrsC) & vbCrLf & NewCaptionTxt("07029", adrsC) & _
            LCode & NewCaptionTxt("07030", adrsC), vbYesNo + vbQuestion)
                Case 7 'no
                    txtFrom.SetFocus
                    ImmediateAbsent = False
                    Exit Function
            End Select
        End If
    End If
End If

strTmp = ""
strPATmp = ""
bytPATmp = 0
strTmp = GetMnlTrnFile((DateCompDate(txtTo.Text) + 1))
If FindTable(strTmp) Then
    If adrsPaid.State = 1 Then adrsPaid.Close
    adrsPaid.Open "Select Presabs from " & strTmp & " where " & strKDate & " =" & strDTEnc & _
    DateCompStr(CStr(DateCompDate(txtTo.Text) + 1)) & strDTEnc & " and Empcode=" & _
    "'" & ECode & "'", VstarDataEnv.cnDJConn
    If Not (adrsPaid.BOF And adrsPaid.EOF) Then
        strPATmp = adrsPaid("Presabs")
        bytPATmp = Len(strPATmp)
        If Mid(strPATmp, 1, bytPATmp / 2) = pVStar.AbsCode Then
            Select Case MsgBox(NewCaptionTxt("07028", adrsC) & vbCrLf & NewCaptionTxt("07029", adrsC) & _
            LCode & NewCaptionTxt("07030", adrsC), vbYesNo + vbQuestion)
                Case 7 'no
                    txtFrom.SetFocus
                    ImmediateAbsent = False
                    Exit Function
            End Select
        End If
    End If
End If
Exit Function
ERR_P:
    ShowError ("ImmediateAbsent :: " & Me.Caption)
    ImmediateAbsent = False
End Function
Private Function ImmediateLeave() As Boolean        '' Checks if Immidiate Leaves are Taken
On Error GoTo ERR_P                                 '' or not
ImmediateLeave = True
Dim strA_R As String, strHFOPT As String
Dim bytCtr As Byte
strA_R = "select * from LvInfo" & Right(pVStar.YearSel, 2) & " where ((" & strDTEnc & _
DateCompStr(CStr(DateCompDate(txtFrom.Text) - 1)) & strDTEnc & " between fromdate and todate ) or (" & _
strDTEnc & DateCompStr(CStr(DateCompDate(txtTo.Text) + 1)) & strDTEnc & _
" between fromdate and todate)) and trcd=4 and Empcode='" & ECode & "'"
If adrsPaid.State = 1 Then adrsPaid.Close
adrsPaid.Open strA_R, VstarDataEnv.cnDJConn
bytCtr = 0
If Not (adrsPaid.EOF And adrsPaid.BOF) Then
        Do While Not adrsPaid.EOF
                If strF = "F " Then bytCtr = bytCtr + 1
                If strT = "T " Then bytCtr = bytCtr + 1
                strHFOPT = Left(adrsPaid("hf_option"), 2)
                If strHFOPT = strF Or strHFOPT = "F " Then bytCtr = bytCtr + 1
                strHFOPT = Right(adrsPaid("HF_Option"), 2)
                If strHFOPT = strT Or strHFOPT = "T " Then bytCtr = bytCtr + 1
                adrsPaid.MoveNext
        Loop
End If
If bytCtr > 0 Then
    If MsgBox(NewCaptionTxt("07031", adrsC), vbYesNo + vbQuestion) = vbYes Then
        ImmediateLeave = True
    Else
        txtFrom.SetFocus
        ImmediateLeave = False
    End If
End If
Exit Function
ERR_P:
    ShowError ("ImmediateLeave :: " & Me.Caption)
    ImmediateLeave = False
End Function

Private Function SaveAddMaster() As Boolean     '' Saves Data in the Leave Infomation File
On Error GoTo ERR_P                             '' and Updates the Balances of the Employee
Dim strTmp As String
strTmp = CStr(Date)
SaveAddMaster = True                            '' in the Leave Balance File
VstarDataEnv.cnDJConn.BeginTrans
'' Insert Information in LvInfo
    VstarDataEnv.cnDJConn.Execute "insert into LvInfo" & Right((pVStar.YearSel), 2) & _
    " (EmpCode,LCode,Fromdate,Todate,Trcd,Days,Lv_Type_rw,Hf_Option,Entrydate)  values" & _
    "('" & ECode & "','" & LCode & "'," & strDTEnc & DateSaveIns(txtFrom.Text) & _
    strDTEnc & "," & strDTEnc & DateSaveIns(txtTo.Text) & strDTEnc & ",4," & _
    TotDays & "," & "'" & strRW & "','" & strHf_Opt & "'," & strDTEnc & DateSaveIns(strTmp) & strDTEnc & ")"
'' Update balance in LvBal
'If (Not blnNoBal) Then
'    VstarDataEnv.cnDJConn.Execute "Update Lvbal" & Right(pVStar.YearSel, 2) & " Set " & LCode & "=" & sngDaysBal & " Where Empcode='" & ECode & "'"
'End If
'' Update Status in the Monthly Transaction File
Call UpdateStatusOnAdd
VstarDataEnv.cnDJConn.CommitTrans
Exit Function
ERR_P:
    SaveAddMaster = False
    ShowError ("SaveAddMaster :: " & Me.Caption)
End Function

Private Sub UpdateStatusOnAdd()         '' Updates the Status in the Monthly Transaction
On Error GoTo ERR_P                     '' File After Leave is Availed by the Employee
Dim bytCnt As Byte, bytCntTmp As Byte   '' Temporary Variables
Dim strTmpTrn As String, strTmpShf As String    '' Temporary Trn and Shf File Variables
Dim dtTemp As Date                      '' Temporary Variables
dtTemp = DateCompDate(txtFrom.Text)

Do While Month(dtTemp) <= Month(DateCompDate(txtTo.Text)) And Year(dtTemp) <= Year(DateCompDate(txtTo.Text))
    strTmpTrn = MakeName(MonthName(Month(dtTemp)), Year(dtTemp), "Trn")
    strTmpShf = MakeName(MonthName(Month(dtTemp)), Year(dtTemp), "shf")
    '' Only If Monthly Trn File Is Found
    If FindTable(strTmpTrn) Then
        If adrsPaid.State = 1 Then adrsPaid.Close
        ''This If Condition Add By
        adrsPaid.Open "Select " & strKDate & " ,Presabs from " & strTmpTrn & " where Empcode= " & "'" & _
        ECode & "'" & " order by " & strKDate & " ", _
        VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic
        ''
        If Not (adrsPaid.BOF And adrsPaid.EOF) Then
            Do
                If adrsPaid!Date > DateCompDate(txtTo.Text) Then Exit Sub
                If adrsPaid("date") >= DateCompDate(txtFrom.Text) And _
                adrsPaid("date") <= DateCompDate(txtTo.Text) Then
                    '' For Date=From date and <>To Date
                    If adrsPaid("date") = DateCompDate(txtFrom.Text) And _
                    adrsPaid("date") <> DateCompDate(txtTo.Text) Then
                        Call CriteriaOneAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                    '' For Date=To date and <> From date
                    If adrsPaid("date") = DateCompDate(txtTo.Text) And _
                    adrsPaid("date") <> DateCompDate(txtFrom.Text) Then
                        Call CriteriaTwoAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                    '' For Date > From date and < To Date
                    If adrsPaid("date") > DateCompDate(txtFrom.Text) And _
                    adrsPaid("date") < DateCompDate(txtTo.Text) Then
                        Call CriteriaThreeAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                    '' For Date=From date and =To date
                    If adrsPaid("date") = DateCompDate(txtFrom.Text) And _
                    adrsPaid("date") = DateCompDate(txtTo.Text) Then
                        Call CriteriaFourAdd(adrsPaid, strTmpTrn, strTmpShf)
                    End If
                End If
                adrsPaid.MoveNext
            Loop Until adrsPaid.EOF
            adrsPaid.Close
        End If
    End If
    dtTemp = DateAdd("m", 1, dtTemp)
Loop
Exit Sub
ERR_P:
    ShowError ("UpdateStatusOnAdd :: " & Me.Caption)
End Sub


Private Function GetMnlTrnFile(ByVal dt As String) As String
On Error GoTo ERR_P
Dim Mon_trn As String
Mon_trn = MakeName(MonthName(Month(CDate(dt))), Year(CDate(dt)), "trn")
If FindTable(Mon_trn) Then
    GetMnlTrnFile = Mon_trn
Else
    GetMnlTrnFile = ""
End If
Exit Function
ERR_P:
    ShowError ("GetMnlTrnFile :: Common ")
    GetMnlTrnFile = ""
End Function

Private Sub CriteriaOneAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case Left(strHf_Opt, (Len(strHf_Opt) / 2))
    Case "FF"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & _
                " where Empcode=" & "'" & ECode & "'" & " and " & strTmpTrn _
                & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & _
            "'" & ECode & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
        
    Case "FS"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, LCode) & _
                "'" & " where Empcode=" & "'" & ECode & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & StuffVal(adrsRec!presabs, 3, 2, LCode) & "'" & " where Empcode=" & _
            "'" & ECode & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
        
    Case "F "
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & _
                "'" & ECode & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
                DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & "'" & ECode & _
            "'" & " and  " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
        
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaOneAdd ::" & Me.Caption)
End Sub

Private Sub CriteriaTwoAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case Right(strHf_Opt, Len(strHf_Opt) / 2)
    Case "TF"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 1, 2, LCode) & "'" & _
                " where Empcode=" & "'" & ECode & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & StuffVal(adrsRec!presabs, 1, 2, LCode) & "'" & " where Empcode=" & _
            "'" & ECode & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
    Case "TS"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & _
                " where Empcode=" & "'" & ECode & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & "'" & ECode & _
            "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
    Case "T "
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & _
                " where Empcode=" & "'" & ECode & "'" & " and " & strTmpTrn & _
                "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & " set Presabs=" & _
            "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & "'" & ECode & "'" & _
            " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaTwoAdd :: " & Me.Caption)
End Sub

Private Sub CriteriaThreeAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
    If strRW = "R" Then
        VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
        " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & _
        "'" & ECode & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
        DateCompStr(adrsRec!Date) & strDTEnc
    End If
Else
    VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & " set Presabs=" & _
    "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & "'" & ECode & _
    "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
End If
Exit Sub
ERR_P:
    ShowError ("CriteriaThreeAdd :: " & Me.Caption)
End Sub

Private Sub CriteriaFourAdd(ByRef adrsRec As ADODB.Recordset, _
ByVal strTmpTrn As String, ByVal strTmpShf As String)
On Error GoTo ERR_P
Select Case strHf_Opt
    Case "FFTF"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 1, 2, LCode) & _
                "'" & " where Empcode=" & "'" & ECode & "'" & " and " & _
                strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 1, 2, LCode) & "'" & _
            " where Empcode=" & "'" & ECode & "'" & " and " & strTmpTrn & _
            "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
    Case "FSTS"
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
                VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, LCode) & _
                "'" & " where Empcode=" & "'" & ECode & "'" & " and " & _
                strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & StuffVal(adrsRec!presabs, 3, 2, LCode) & _
            "'" & " where Empcode=" & "'" & ECode & "'" & " and " & _
            strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
        End If
    Case "F T "
        If Left(adrsRec("Presabs"), 2) = pVStar.WosCode Or _
        Left(adrsRec("Presabs"), 2) = pVStar.HlsCode Then
            If strRW = "R" Then
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
                " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & _
                " where Empcode=" & "'" & ECode & "'" & " and " & _
                strTmpTrn & "." & strKDate & " =" & strDTEnc & DateCompStr(adrsRec!Date) & strDTEnc
            End If
        Else
            VstarDataEnv.cnDJConn.Execute "update " & strTmpTrn & _
            " set Presabs=" & "'" & ReplicateVal(LCode, 2) & "'" & " where Empcode=" & _
            "'" & ECode & "'" & " and " & strTmpTrn & "." & strKDate & " =" & strDTEnc & _
            DateCompStr(adrsRec!Date) & strDTEnc
        End If
End Select
Exit Sub
ERR_P:
    ShowError ("CriteriaFourAdd :: " & Me.Caption)
End Sub

Public Function ZeroPadding(L As Integer, strTemp As String) As String
Dim i As Integer
For i = 1 To L
    strTemp = "0" & strTemp
Next
ZeroPadding = strTemp
End Function

Private Function ImportData(FilePath As String) As Boolean
On Error GoTo ERR_P
Dim strSql As String
Dim lng As Long
Dim bytTmp As Integer, i As Integer, Flag As Integer
Dim rsImport As New ADODB.Recordset
Dim rsSftCode As New ADODB.Recordset
Dim rsLvCode As New ADODB.Recordset
Dim xlSheet As Excel.Worksheet
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim name As String, Designatn As String, Empcode As String, card As String
Dim joindate As String, confmdt As String, birth_dt As String, leavdate As String, shf_date As String
Dim div As Integer, dept As Integer, cat As Integer, entry As Integer, company As Integer, Location As Integer, MachineCode As Integer
Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    xlApp.Workbooks.Open CD1.FileName
    xlApp.Workbooks(1).Activate
    Set xlSheet = xlApp.ActiveWorkbook.Sheets(1)
If FilePath = "" Then
    MsgBox "No file has been selected.", vbInformation
    Exit Function
Else
    If Right(FilePath, 4) = ".xls" Then
        If FindTable("ImportTbl") Then VstarDataEnv.cnDJConn.Execute "drop table ImportTbl"
        Call CreateTableIntoAs("*", "empmst", "ImportTbl", " where 1=2")

            Dim cn As New ADODB.Connection
            'Set cn = New ADODB.Connection
            Dim rsTemp As New ADODB.Recordset
            Dim rsAccess As New ADODB.Recordset
            If cn.State = 1 Then cn.Close
            cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
            cn.Open
            VstarDataEnv.cnDJConn.Execute "delete from ImportTbl"
            'VstarDataEnv.cnDJConn.Execute "drop table ImportTbl"
            strSql = "select * FROM [Excel 8.0;DATABASE=" & FilePath & ";HDR=Yes].[sheet1$]"
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open strSql, cn, adOpenStatic
         
         If rsAccess.State = 1 Then rsAccess.Close
        rsAccess.Open "select * from ImportTbl", VstarDataEnv.cnDJConn, adOpenDynamic, adLockOptimistic
        For i = 1 To rsTemp.RecordCount 'looping through all records of the excel file
            'Empcode = Trim(xlSheet.Cells(i + 1, 1).Value)
            'If Not (ECode = "") Then
                'If Len(ECode) <> ECodeSize Then ECode = ZeroPadding(ECodeSize - Len(ECode), ECode)
                name = xlSheet.Cells(i + 1, 1).Value
                Designatn = xlSheet.Cells(i + 1, 2).Value
                Empcode = xlSheet.Cells(i + 1, 3).Value
                card = xlSheet.Cells(i + 1, 4).Value
                joindate = xlSheet.Cells(i + 1, 5).Value
                confmdt = xlSheet.Cells(i + 1, 6).Value
                div = xlSheet.Cells(i + 1, 7).Value
                entry = xlSheet.Cells(i + 1, 8).Value
                birth_dt = xlSheet.Cells(i + 1, 9).Value
                leavdate = xlSheet.Cells(i + 1, 10).Value
                shf_date = xlSheet.Cells(i + 1, 11).Value
                MachineCode = xlSheet.Cells(i + 1, 12).Value
                
                rsAccess.AddNew
                rsAccess.Fields("empcode") = Empcode
                rsAccess.Fields("name") = name
                rsAccess.Fields("Designatn") = Designatn
                rsAccess.Fields("empcode") = Empcode
                rsAccess.Fields("card") = card
                rsAccess.Fields("joindate") = joindate
                rsAccess.Fields("confmdt") = confmdt
                rsAccess.Fields("div") = div
                rsAccess.Fields("entry") = entry
                rsAccess.Fields("birth_dt") = birth_dt
                rsAccess.Fields("leavdate") = leavdate
                rsAccess.Fields("shf_date") = shf_date
                rsAccess.Fields("MachineCode") = MachineCode
            rsAccess.Update
            'End If
            Empcode = "": name = "": Designatn = "": Empcode = "": card = "": div = 0: dept = 0: cat = 0: entry = 0
            company = 0: Location = 0: MachineCode = 0
            rsTemp.MoveNext
        Next i
        'End Select
               
        If rsImport.State = 1 Then rsImport.Close
        rsImport.Open "select * from ImportTbl", VstarDataEnv.cnDJConn, adOpenStatic
        
        If Not (rsImport.EOF And rsImport.BOF) Then
            Screen.MousePointer = vbHourglass
            rsImport.MoveFirst
            While Not (rsImport.EOF)
                VstarDataEnv.cnDJConn.Execute "insert into EmpMst(EmpCode,Card,Name,Designatn,Entry,joindate," & _
                "confmdt,birth_dt,leavdate,Div,MachineCode,shf_date) values('" & rsAccess.Fields("empcode") & "','" & rsAccess.Fields("card") & _
                "','" & rsAccess.Fields("name") & "','" & rsAccess.Fields("Designatn") & "','" & rsAccess.Fields("Entry") & "','" & rsAccess.Fields("joindate") & "','" & rsAccess.Fields("confmdt") & _
                "','" & rsAccess.Fields("birth_dt") & "','" & rsAccess.Fields("leavdate") & "'," & rsAccess.Fields("Div") & "," & rsAccess.Fields("MachineCode") & ",'" & rsAccess.Fields("shf_date") & "')"
                rsImport.MoveNext
            Wend
        
            Screen.MousePointer = vbNormal
            MsgBox "Import sucessfully."
         End If
    Else
         MsgBox "Selected File is not excel file. Select excel file and Try Again.", vbOKOnly + vbInformation, "Imoport Excel File"
         Exit Function
    End If
End If
Exit Function
ERR_P:
'MsgBox Err.Number & " " & Err.Description
If Err.Number = -2147217900 Then
Resume Next
    MsgBox "Excel file is in edit mode. First close the excel file and then import.", vbOKOnly + vbInformation, "Imoport Excel File"
'    MsgBox Err.Description
Else
    ShowError ("ImportData :: " & Me.Caption)
End If
End Function

