VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optSign 
      Caption         =   "Deduct"
      Height          =   195
      Index           =   1
      Left            =   3375
      TabIndex        =   10
      Top             =   840
      Width           =   960
   End
   Begin VB.OptionButton optSign 
      Caption         =   "Add"
      Height          =   195
      Index           =   0
      Left            =   2655
      TabIndex        =   9
      Top             =   840
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.TextBox txtADJ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4365
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "D"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame frDates 
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5805
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Tag             =   "D"
         Top             =   210
         Width           =   1575
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Tag             =   "D"
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label lblFromD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&From Date"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblToD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&To Date"
         Height          =   195
         Left            =   3180
         TabIndex        =   2
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdEmpMaster 
      Caption         =   "EmpMaster"
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "&Export"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Previous month adjustments (if any)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adrsEX As New ADODB.Recordset
Dim dtFrom As Date, dtTo As Date
Dim strF1 As String, strF2 As String, strLvFile As String
Dim sngADJ As Single

Private Sub cmdEmpMaster_Click()
        Screen.MousePointer = vbHourglass
        cmdEmpMaster.Enabled = False
        cmdExit.Enabled = False
        Dim strqry As String
        strqry = "select *from Empmst order by empcode"
            Screen.MousePointer = vbHourglass
            frmExp.SaveExlExport (strqry)
            Call ChkRepFile
            cmdEmpMaster.Enabled = True:  cmdExit.Enabled = True
            Screen.MousePointer = vbNormal
    Exit Sub
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdExp_Click()
'' Get the File Name
If Not CheckDates Then Exit Sub

''Get Adjustment
sngADJ = IIf(Trim(txtADJ.Text) = "", 0, Val(Trim(txtADJ.Text)))
If optSign(1).Value Then sngADJ = sngADJ * -1
Screen.MousePointer = vbHourglass
If Exported Then
    MsgBox "Export complete", vbInformation
End If
Screen.MousePointer = vbNormal
End Sub

Private Function CheckDates() As Boolean    '' Function to Check if Dates are in Valid Range
On Error GoTo ERR_P
Dim strDateM
If Trim(txtFrom.Text) = "" Then
    MsgBox "Please Enter From Date", vbExclamation
    CheckDates = False
    txtFrom.SetFocus
    Exit Function
End If
If Trim(txtTo.Text) = "" Then
    MsgBox "Please Enter To Date", vbExclamation
    CheckDates = False
    txtTo.SetFocus
    Exit Function
End If
dtFrom = DateCompDate(txtFrom.Text)
dtTo = DateCompDate(txtTo.Text)
strF1 = MonthName(Month(dtFrom), True) & Right(CStr(Year(dtFrom)), 2) & "Trn"
strF2 = MonthName(Month(dtTo), True) & Right(CStr(Year(dtTo)), 2) & "Trn"
strLvFile = "LVBAL" & Right(Year(dtTo), 2)

If Not FindTable(strF1) Then
    MsgBox "Transaction file for " & MonthName(Month(dtFrom)) & " not found.", vbInformation
    Exit Function
End If
If Not FindTable(strF2) Then
    MsgBox "Transaction file for " & MonthName(Month(dtTo)) & " not found.", vbInformation
    Exit Function
End If
CheckDates = True
Exit Function
ERR_P:
    ShowError ("Check Dates :: " & Me.Caption)
End Function

Private Function Exported() As Boolean
On Error GoTo ERR_P
Dim strQ As String, STRECODE As String, Present As String
Dim strquery As String, strInsert As String
Dim sngDaysworked As Single, sngMedical As Single
Dim sngOD As Single, sngCO As Single
Dim sngCL As Single, sngSL As Single
Dim sngPL As Single, sngHL As Single
Dim sngLAYOFF As Single
Dim sngWK As Single, sngAL As Single
Dim sngAA As Single, sngSUS As Single
Dim sngTOTPAYDAY As Single, sngBALCL As Single
Dim sngBALSL As Single, sngBALPL As Single

strInsert = "Insert into ATTENDANCE(TOKENNO, MMMYY, DAYSWORKED, MEDICAL, OD,COMPOFF, CL," & _
" SL,PL,PH,ADJUSTMENTS, LAYOFF, WK, PAYAUTHORISED, PAYUNAUTHORISED, SUSPN, TOTALDAYSPAYABLE, " & _
" BALCL, BALSL, BALPL,PresentDays) values("

If strF1 = strF2 Then
    strQ = "Select EMPMST.empcode,[date],presabs,ovtim,otconf,wrkhrs,stfwrk,present from " & _
    strF1 & ",Empmst where " & strF1 & ".Empcode = Empmst.Empcode  " & _
    " and [date] >= " & strDTEnc & DateCompStr(dtFrom) & strDTEnc & " and [date] <= " & _
    strDTEnc & DateCompStr(dtTo) & strDTEnc & " order by EMPMST.empcode,[date]"
Else
    strQ = "Select EMPMST.empcode,[date],presabs,ovtim,otconf,wrkhrs from " & _
    strF1 & ",Empmst where " & strF1 & ".Empcode = Empmst.Empcode " & _
    " and [date] >= " & strDTEnc & DateCompStr(dtFrom) & strDTEnc & " union " & _
    " Select EMPMST.empcode,[date],presabs,ovtim,otconf,wrkhrs from " & _
    strF2 & ",Empmst where " & strF2 & ".Empcode = Empmst.Empcode " & _
    " and [date] <= " & strDTEnc & DateCompStr(dtTo) & strDTEnc & " order by EMPMST.empcode,[date]"
End If

''cleaning the ATTENDANCE TABLE
VstarDataEnv.cnDJConn.Execute "DELETE from ATTENDANCE"

If adrsEX.State = 1 Then adrsEX.Close
adrsEX.Open strQ, VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
If Not (adrsEX.EOF And adrsEX.BOF) Then
    Do While Not adrsEX.EOF
        STRECODE = adrsEX("Empcode")
        Present = adrsEX("present")
        Do While STRECODE = adrsEX("Empcode")
            Select Case Left(adrsEX("presabs"), 2)
                Case pVStar.PrsCode
                    sngDaysworked = sngDaysworked + 0.5
                Case "ML"
                    sngMedical = sngMedical + 0.5
                Case "TL"
                    sngOD = sngOD + 0.5
                Case "OD"                                 ''Add Case OD By  15-11-08
                    sngOD = sngOD + 0.5
                Case "CO"
                    sngCO = sngCO + 0.5
                Case "CL"
                    sngCL = sngCL + 0.5
                Case "SL"
                    sngSL = sngSL + 0.5
                Case "PL"
                    sngPL = sngPL + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case "LY"
                    sngLAYOFF = sngLAYOFF + 0.5
                Case pVStar.WosCode
                    If adrsEX("STFWRK") = "S" Then
                        sngWK = sngWK + 0.5
                    End If
                Case "AL"
                    sngAL = sngAL + 0.5
                Case "AA"
                    sngAA = sngAA + 0.5
                Case "SU"
                    sngSUS = sngSUS + 0.5
            End Select
            Select Case Right(adrsEX("presabs"), 2)
                Case pVStar.PrsCode
                    sngDaysworked = sngDaysworked + 0.5
                Case "ML"
                    sngMedical = sngMedical + 0.5
                Case "TL"
                    sngOD = sngOD + 0.5
                Case "OD"                           ''Add Case OD By  15-11-08
                    sngOD = sngOD + 0.5
                Case "CO"
                    sngCO = sngCO + 0.5
                Case "CL"
                    sngCL = sngCL + 0.5
                Case "SL"
                    sngSL = sngSL + 0.5
                Case "PL"
                    sngPL = sngPL + 0.5
                Case pVStar.HlsCode
                    sngHL = sngHL + 0.5
                Case "LY"
                    sngLAYOFF = sngLAYOFF + 0.5
                Case pVStar.WosCode
                    If adrsEX("STFWRK") = "S" Then
                        sngWK = sngWK + 0.5
                    End If
                Case "AL"
                    sngAL = sngAL + 0.5
                Case "AA"
                    sngAA = sngAA + 0.5
                Case "SU"
                    sngSUS = sngSUS + 0.5
            End Select

            adrsEX.MoveNext
            If adrsEX.EOF Then Exit Do
        Loop
        
        sngTOTPAYDAY = sngDaysworked + (sngMedical / 2) + sngOD + sngCO + sngCL + sngSL + sngPL + sngHL + sngADJ + sngWK + sngLAYOFF
        
        If adrsTemp.State = 1 Then adrsTemp.Close
        adrsTemp.Open "select CL, SL, PL from " & strLvFile & " where empcode = '" & STRECODE & "'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
        If Not (adrsTemp.EOF And adrsTemp.BOF) Then
            sngBALCL = IIf(IsNull(adrsTemp("CL")), 0, adrsTemp("CL"))
            sngBALSL = IIf(IsNull(adrsTemp("SL")), 0, adrsTemp("SL"))
            sngBALPL = IIf(IsNull(adrsTemp("PL")), 0, adrsTemp("PL"))
        End If
        
        strInsert = strInsert & "'" & Val(STRECODE) & "','" & MonthName(Month(dtTo), True) & Right(Year(dtTo), 2) & "'," & _
        sngDaysworked & ", " & sngMedical & ", " & sngOD & "," & sngCO & "," & sngCL & "," & sngSL & "," & _
        sngPL & "," & sngHL & "," & sngADJ & "," & sngLAYOFF & "," & sngWK & "," & sngAL & "," & _
        sngAA & "," & sngSUS & "," & sngTOTPAYDAY & "," & sngBALCL & "," & sngBALSL & "," & sngBALPL & "," & Present & ")"
        
        VstarDataEnv.cnDJConn.Execute strInsert
        
        sngDaysworked = 0: sngMedical = 0: sngOD = 0: sngCO = 0
        sngCL = 0: sngSL = 0: sngPL = 0: sngHL = 0
        sngLAYOFF = 0: sngWK = 0: sngAL = 0
        sngAA = 0: sngSUS = 0: sngTOTPAYDAY = 0: sngBALCL = 0
        sngBALSL = 0: sngBALPL = 0
        
        strInsert = "Insert into ATTENDANCE(TOKENNO, MMMYY, DAYSWORKED, MEDICAL, OD,COMPOFF, CL," & _
        " SL,PL,PH,ADJUSTMENTS, LAYOFF, WK, PAYAUTHORISED, PAYUNAUTHORISED, SUSPN, TOTALDAYSPAYABLE, " & _
        " BALCL,BALSL,BALPL,PresentDays) values("

    Loop
Else
    MsgBox "No Records Found.", vbInformation
    Exit Function
End If
Exported = True
Exit Function
ERR_P:
    ShowError ("Exported :: " & Me.Caption)
    'Resume Next
End Function

Private Sub Form_Load()
Call SetFormIcon(Me)        '' Set the Form Icon
End Sub

Private Sub txtADJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
Else
    KeyAscii = KeyPressCheck(KeyAscii, 2)
End If
End Sub

Private Sub txtFrom_Click()
varCalDt = ""
varCalDt = Trim(txtFrom.Text)
txtFrom.Text = ""
Call ShowCalendar
End Sub

Private Sub txtFrom_GotFocus()
    Call GF(txtFrom)
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    Call CDK(txtFrom, KeyAscii)
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
If Not ValidDate(txtFrom) Then txtFrom.SetFocus: Cancel = True
End Sub

Private Sub txtTo_Click()
varCalDt = ""
varCalDt = Trim(txtTo.Text)
txtTo.Text = ""
Call ShowCalendar
End Sub

Private Sub txtTo_GotFocus()
    Call GF(txtTo)
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    Call CDK(txtTo, KeyAscii)
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
    If Not ValidDate(txtTo) Then txtTo.SetFocus: Cancel = True
End Sub

