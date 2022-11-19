VERSION 5.00
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   " "
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   930
   ClientWidth     =   7680
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   Picture         =   "MainForm.frx":2EFA
   ScrollBars      =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picWin 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   7620
      TabIndex        =   0
      Top             =   0
      Width           =   7680
      Begin VB.Image ImgWin 
         Height          =   735
         Left            =   480
         Stretch         =   -1  'True
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Menu Mnusetup 
      Caption         =   "Main Setting"
      Begin VB.Menu ParameterMnu 
         Caption         =   "Parameter"
      End
      Begin VB.Menu mnuOTRules 
         Caption         =   "O&T Rules"
      End
      Begin VB.Menu mnuCORules 
         Caption         =   "C&O Rules"
      End
      Begin VB.Menu mnulaterule 
         Caption         =   "MonthlyLateRule"
      End
      Begin VB.Menu RuleMnu 
         Caption         =   "Late / Early Rule"
      End
      Begin VB.Menu SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiff 
         Caption         =   "Login as &Different User"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu ExtMnu 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "Master"
      Begin VB.Menu mnuCompany 
         Caption         =   "Company"
      End
      Begin VB.Menu ShftMnu 
         Caption         =   "Shift"
      End
      Begin VB.Menu CatMastMnu 
         Caption         =   "Category"
      End
      Begin VB.Menu LvMnu 
         Caption         =   "Leave"
      End
      Begin VB.Menu DepMastMnu 
         Caption         =   "Department"
      End
      Begin VB.Menu grpMnu 
         Caption         =   "Group Master"
      End
      Begin VB.Menu grademnu 
         Caption         =   "Grade Master"
      End
      Begin VB.Menu RotMastMnu 
         Caption         =   "Rotation"
      End
      Begin VB.Menu mnuDiv 
         Caption         =   "&Division Master"
      End
      Begin VB.Menu mnuLoca 
         Caption         =   "Location Master"
      End
      Begin VB.Menu mnuDesigMst 
         Caption         =   "Designation Master"
      End
      Begin VB.Menu HDayMnu 
         Caption         =   "Holiday"
         Begin VB.Menu HolMastMnu 
            Caption         =   "Holiday Master"
         End
         Begin VB.Menu DecHdayMnu 
            Caption         =   "Declare Holiday"
         End
      End
   End
   Begin VB.Menu updatemnu 
      Caption         =   "Updation"
      Begin VB.Menu EmpMastMnu 
         Caption         =   "Employee"
      End
      Begin VB.Menu Mnuset 
         Caption         =   "&Set Employee Details"
      End
      Begin VB.Menu mnuShiftCh 
         Caption         =   "Change Schedule for All"
      End
      Begin VB.Menu mnuOTAuth 
         Caption         =   "OT &Authorization"
      End
      Begin VB.Menu CorrectMnu 
         Caption         =   "Correction"
      End
      Begin VB.Menu EdiPaidMnu 
         Caption         =   "&Edit Paid days"
      End
      Begin VB.Menu mnuLost 
         Caption         =   "Manual Entry"
      End
   End
   Begin VB.Menu MnuLabour 
      Caption         =   "Work Planning"
      Begin VB.Menu SchMastMnu 
         Caption         =   "Shedule Master"
      End
      Begin VB.Menu CrSchMnu 
         Caption         =   "Create Planning"
      End
      Begin VB.Menu ChgSchMnu 
         Caption         =   "Change Shedule"
      End
   End
   Begin VB.Menu ProcessMnu 
      Caption         =   "Process"
      Begin VB.Menu DailyPMnu 
         Caption         =   "Daily Process"
      End
      Begin VB.Menu MthlyPMnu 
         Caption         =   "Monthly Process"
      End
      Begin VB.Menu YrLvMnu 
         Caption         =   "Yearly"
         Begin VB.Menu YrCrtMnu 
            Caption         =   "Yearly Leave Create"
         End
         Begin VB.Menu YrUptMnu 
            Caption         =   "Yearly Leave Update"
         End
      End
   End
   Begin VB.Menu MnuLeave 
      Caption         =   "Leave Request"
      Begin VB.Menu OpenLvMnu 
         Caption         =   "Opening"
      End
      Begin VB.Menu CrdtLvMnu 
         Caption         =   "Credit"
      End
      Begin VB.Menu EncLvMnu 
         Caption         =   "Encash"
      End
      Begin VB.Menu AvlLvMnu 
         Caption         =   "Avail"
      End
   End
   Begin VB.Menu ReportMnu 
      Caption         =   "Reports"
      Begin VB.Menu RprtMnu 
         Caption         =   "Reports"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDailyReports 
         Caption         =   "Daily"
      End
      Begin VB.Menu mnuWeeklyReports 
         Caption         =   "Weekly"
      End
      Begin VB.Menu mnuMonthlyReports 
         Caption         =   "Monthly"
      End
      Begin VB.Menu mnuYearReport 
         Caption         =   "Yearly"
      End
      Begin VB.Menu mnuBetweenDatere 
         Caption         =   "Periodic"
      End
      Begin VB.Menu mnuMasterReport 
         Caption         =   "Master"
      End
      Begin VB.Menu mnuAudit 
         Caption         =   "Audit Report"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu INI 
      Caption         =   "INI"
      Visible         =   0   'False
   End
   Begin VB.Menu UtilityMnu 
      Caption         =   "Utility"
      Begin VB.Menu mnuAdmin 
         Caption         =   "Admin &Form"
      End
      Begin VB.Menu usrgrpMnu 
         Caption         =   "User and Group &Account"
      End
      Begin VB.Menu mnuChUPass 
         Caption         =   "C&hange User Password"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainExp 
         Caption         =   "Export"
         Begin VB.Menu mnuExport 
            Caption         =   "&Export Data"
         End
         Begin VB.Menu mnuPerformance 
            Caption         =   "Performance Data"
         End
         Begin VB.Menu mnuMonthlyEntries 
            Caption         =   "Monthly Entries"
         End
         Begin VB.Menu mnuExportCustom 
            Caption         =   "Export Custom"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuImportLogs 
         Caption         =   "Import Logs"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusep12 
         Caption         =   "-"
      End
      Begin VB.Menu VMnu 
         Caption         =   "Version"
      End
   End
   Begin VB.Menu SupportMnu 
      Caption         =   " support"
      Begin VB.Menu mnuOnline 
         Caption         =   "&Online Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuRC 
      Caption         =   "Right Click"
      Visible         =   0   'False
      Begin VB.Menu mnuRCSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRCCorr 
         Caption         =   "&Correction"
      End
      Begin VB.Menu mnuRCDaily 
         Caption         =   "&Daily Process"
      End
      Begin VB.Menu mnuRCMonthly 
         Caption         =   "&Monthly Process"
      End
      Begin VB.Menu SEP10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRAvail 
         Caption         =   "&Leave Avail"
      End
      Begin VB.Menu mnuRCSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCChgSch 
         Caption         =   "&Change Schedule"
      End
      Begin VB.Menu mnuRCSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCLogin 
         Caption         =   "&Login as Different User"
      End
      Begin VB.Menu mnuRCSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnRemind As Boolean
Dim adrsC As New ADODB.Recordset

Private Sub AvlLvMnu_Click()
bytFormToLoad = 4       '' Avail Leaves
PassFrm.Show vbModal
End Sub

Private Sub EdiPaidMnu_Click()
    frmEditPaid.Show vbModal
End Sub

Private Sub grademnu_Click()
    frmGrade.Show vbModal
End Sub

Private Sub grpMnu_Click()
    frmGroup.Show vbModal
End Sub

Private Sub ImgWin_DblClick()
Static intTmp As Integer
Static bytMin As Byte
intTmp = intTmp + 1
If intTmp = 1 Then bytMin = Minute(Now)
If intTmp = 5 Then
    If bytMin = Minute(Now) Then
        If UCase(UserName) = UCase(strPrintUser) Then
            bytFormToLoad = 6
            PassFrm.Show vbModal
        End If
    End If
    intTmp = 0
    bytMin = 0
End If
End Sub

Private Sub ImgWin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    Call PopupMenu(mnuRC, vbPopupMenuLeftAlign, x, y, mnuRCCorr)
End If
End Sub


Private Sub INI_Click()
frmPE.Show
End Sub

Private Sub MDIForm_Activate()
On Error GoTo ERR_P
If strCurrentUserType <> HOD Then
    mnuAdmin.Visible = True
Else
    mnuAdmin.Visible = False
End If
If blnDBCompacted = True Then Exit Sub
With MainForm
    .Mnusetup.Visible = True
    .updatemnu.Visible = True
    .ProcessMnu.Visible = True
    .ReportMnu.Visible = True
    .UtilityMnu.Visible = True
    ''Install Menu

End With


    MainForm.grademnu.Visible = False


If GetFlagStatus("DeviceLog") Then
    mnuImportLogs.Visible = True
    If Not FindTable("DeviceLog") Then
        ConMain.Execute "Create Table DeviceLog (CardCode text (10), PDate DateTime)"
    End If
End If

If GetFlagStatus("ExportWorks") Then mnuExportCustom.Visible = True

Exit Sub
ERR_P:
    ShowError ("Application Activate Error :: MDI")
    'Resume Next
End Sub

Private Sub MDIForm_Click()
frmExpTimeCard.Show
End Sub

Private Sub MDIForm_DblClick()
    frmExpTimeCard.Show
End Sub

Private Sub MDIForm_Load()
On Error GoTo Err_particular
blnRemind = True

With ImgWin
    .Picture = LoadPicture(App.path & "\Images\TimeHR.gif")
End With
If bytBackEnd = 1 Or bytBackEnd = 3 Then
    blnBackRes = False      '' SQL Server/Oracle
Else
    blnBackRes = True       '' Access
End If
Call SetFormIcon(Me)


    EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
    
     TB_WindowItemSize Me, mntresREs
'*****************************************************
If adRsInstall.State = 1 Then adRsInstall.Close
adRsInstall.Open "select * from install", ConMain, adOpenKeyset, adLockOptimistic
With pVStar
    .CodeSize = IIf(IsNull(adRsInstall("e_codesize")), 0, adRsInstall("e_codesize"))
    .CardSize = IIf(IsNull(adRsInstall("e_cardsize")), 0, adRsInstall("e_cardsize"))
    .Yearstart = IIf(IsNull(adRsInstall("yearfrom")), "", adRsInstall("yearfrom"))
    .YearSel = IIf(IsNull(adRsInstall("cur_year")), "", adRsInstall("cur_year"))
    .WeekStart = IIf(IsNull(adRsInstall("weekfrom")), "", adRsInstall("weekfrom"))
    .Use_Mail = IIf(adRsInstall("email"), True, False)
    .Cust_code = IIf(adRsInstall("definod"), False, adRsInstall("definod"))
End With
Call GetPvstarVars
'' READING COMPANY NAME FROM COMPANY TABLE
Call SetCaptionMainForm
ConMain.Execute "Update Exc set UserNumber=UserNumber+1"
Call RetUserNumber

''
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from newcaptions where ID like '53%'", ConMain, adOpenStatic

Call SetMenu
Exit Sub
Err_particular:
    ShowError ("Application Load Error :: MDI")
End Sub
Private Sub SetMenu()

    mnuOnline.Visible = False
    mnulaterule.Visible = False
   
End Sub
Private Sub HolMastMnu_Click()
    frmHoliday.Show vbModal
End Sub

Private Sub LvMnu_Click()
frmLeaves.Show vbModal
YRCRFrm.Show vbModal
End Sub


Private Sub MDIForm_Resize()
On Error GoTo ERR_P
With picWin
    .Top = Me.Top
    .Height = Me.Height
End With
Exit Sub
ERR_P:
    ShowError ("Application Resize Error :: MDI")
End Sub

Public Sub MDIForm_Unload(Cancel As Integer)
'MsgBox " UNLOAD"
On Error GoTo ERR_P
ConLog.Execute "Update RECNUM Set RecNumber=" & typLog.lngRecord
ConLog.Execute "Update TransDet set TransType = " & Msr_no & ", TransDesc='F' Where TransSource=35"
  
If ConMain.State = 0 Then _
    ConMain.Open
''
If InVar.strSer = 3 Then
'  Dated 28/03/06
   ConMain.Execute "Delete  from errorlog where errorid=" & Msr_no
Else
   ConMain.Execute "Delete from errorlog where errorid=" & Msr_no
End If
End
Exit Sub
ERR_P:
    ShowError ("Application Unload Error :: MDI")
    End
End Sub

Private Sub mnuAbout_Click()
On Error GoTo ERR_P


Dim SuppHlp As New InternetExplorer
SuppHlp.Visible = True
Call SuppHlp.Navigate(App.path & "\Data\helptxt.htm")
Exit Sub
ERR_P:
    ShowError ("Support :: " & Me.Caption)
End Sub

Private Sub mnuAdmin_Click()
    frmAdmin.Show vbModal
End Sub

Private Sub mnuAudit_Click()
frmAuditInfo.Show
End Sub


Private Sub mnuBetweenDatere_Click()
    bytRepMode = 6
    ReportType = "Periodic"
    frmReports.Caption = "Periodic Reports"
    frmReports.Show vbModal
End Sub

Private Sub mnuChUPass_Click()
frmChgPass.Show vbModal
'frmPass.Show vbModal
End Sub

Private Sub mnuCompany_Click()
    frm_mst_Company.Show vbModal
End Sub

Private Sub mnuCORules_Click()
frmCORul.Show vbModal
End Sub

Private Sub mnuDailyReports_Click()
    bytRepMode = 1
    ReportType = "Daily"
    frmReports.Caption = "Daily Reports"
    frmReports.cboSelectReport.Text = "Performance"
    frmReports.Show vbModal
   
End Sub

Private Sub mnuDesigMst_Click()
frmDesg.Show vbModal
End Sub

Private Sub mnuDiff_Click()
LoginStatus = True
frmLogin.Show vbModal
Call SetCaptionMainForm
End Sub

Private Sub mnuDiv_Click()
frmDiv.Show vbModal
End Sub

Private Sub mnuExport_Click()
    'frmExpTimeCard.Show vbModal
    frmExp.Show vbModal
End Sub

Private Sub mnuExportCustom_Click()
    frmExportCustom.Show vbModal
End Sub

Private Sub mnuImportLogs_Click()
    frmImport.Show vbModal
End Sub

Private Sub mnuLoca_Click()
frmLoca.Show vbModal
End Sub

Private Sub mnuLost_Click()
    bytFormToLoad = 9       '' Lost Entry
    PassFrm.Show vbModal
   ' frmLostN.Show vbModal
End Sub

Private Sub mnuMasterReport_Click()
    bytRepMode = 5
    ReportType = "Masters"
    frmReports.Caption = "Master Reports"
    frmReports.Show vbModal
End Sub

Private Sub mnuMonthlyEntries_Click()
    strRepName = "MonEntries"
    frmExportReport.Show vbModal
End Sub

Private Sub mnuMonthlyReports_Click()
    bytRepMode = 3
    ReportType = "Monthly"
    frmReports.Caption = "Monthly Reports"
    frmReports.Show vbModal
End Sub

Private Sub mnuOnline_Click()
    Shell App.path & "\client.exe", vbNormalFocus
End Sub

Private Sub mnuOTAuth_Click()
frmOT.Show vbModal
End Sub

Private Sub mnuOTRules_Click()
frmOTRul.Show vbModal
End Sub

Private Sub mnuPerformance_Click()
    strRepName = "Performance"
    frmExportReport.Show vbModal
End Sub


Private Sub mnuRAvail_Click()
Call AvlLvMnu_Click
End Sub

Private Sub mnuRCCorr_Click()
Call CorrectMnu_Click
End Sub

Private Sub mnuRCChgSch_Click()
    Call ChgSchMnu_Click
End Sub

Private Sub mnuRCDaily_Click()
Call DailyPMnu_Click
End Sub

Private Sub mnuRCExit_Click()
Call ExtMnu_Click
End Sub

Private Sub mnuRCLogin_Click()
Call mnuDiff_Click
End Sub

Private Sub mnuRCMonthly_Click()
Call MthlyPMnu_Click
End Sub

Private Sub mnuSet_Click()
frmSet.Show vbModal
End Sub

Private Sub mnuShiftCh_Click()
    frmShfCh.Show vbModal
End Sub


Private Sub mnuWeeklyReports_Click()
    bytRepMode = 2
    ReportType = "Weekly"
    frmReports.Caption = "Weekly Reports"
    frmReports.Show vbModal
End Sub

Private Sub mnuYearReport_Click()
    bytRepMode = 4
     ReportType = "Yearly"
    frmReports.Caption = "Yearly Reports"
    frmReports.Show vbModal
End Sub

Public Sub MthlyPMnu_Click()
On Error GoTo ERR_P
If adrsRits.State = 1 Then adrsRits.Close
adrsRits.Open "Select * from Exc", ConMain, adOpenKeyset, adLockOptimistic
If adrsRits("Monthly") = 1 Then
    adrsRits.Close
    MsgBox NewCaptionTxt("53045", adrsC), vbExclamation
Else
    adrsRits("Monthly") = 1
    adrsRits.Update
    adrsRits.Close
    frmMonthly.Show vbModal
    Call SetMonthlyFlag
End If
Exit Sub
ERR_P:
    ShowError ("Monthly Process :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub OpenLvMnu_Click()
bytFormToLoad = 1       '' Open Leaves
PassFrm.Show vbModal
End Sub

Private Sub ParameterMnu_Click()
   ParaFrm.Show vbModal
End Sub

Private Sub picWin_Resize()
ImgWin.Top = picWin.Top
ImgWin.Left = picWin.Left
ImgWin.Height = picWin.Height
ImgWin.Width = picWin.Width
End Sub

Private Sub RotMastMnu_Click()
    frmRotation.Show vbModal
End Sub

Private Sub RuleMnu_Click()
    frmRules.Show vbModal
End Sub

Private Sub SchMastMnu_Click()
    bytShfMode = 2
    frmEmpShift.Show vbModal
End Sub

Private Sub ShftMnu_Click()
    frm_mst_Shift.Show vbModal
End Sub

Private Sub usrgrpMnu_Click()
    'UsersFrm.Show vbModal
    frmUser.Show vbModal
End Sub

Private Sub VMnu_Click()
4    Version.Show vbModal
'    frmExpTimeCard.Show vbModal
End Sub

Private Sub YrCrtMnu_Click()
    YRCRFrm.Show vbModal
End Sub

Private Sub YrUptMnu_Click()
On Error GoTo ERR_P
If adrsRits.State = 1 Then adrsRits.Close
    adrsRits.Open "Select * from Exc", ConMain
If adrsRits("yearly") = 1 Then
    MsgBox NewCaptionTxt("53046", adrsC), vbExclamation
Else
    Call SetYearlyFlag(1)
    frmLvUp.Show vbModal
    Call SetYearlyFlag(0)
End If
Exit Sub
ERR_P:
    ShowError ("Yearly Update ::" & Me.Caption)
End Sub

Private Sub CatMastMnu_Click()
    frmCat.Show vbModal
End Sub

Private Sub ChgSchMnu_Click()
    frmSch.Show vbModal
End Sub

Private Sub CorrectMnu_Click()
bytFormToLoad = 5       '' Correction
PassFrm.Show vbModal
End Sub

Private Sub CrdtLvMnu_Click()
bytFormToLoad = 2       '' Credit Leaves
PassFrm.Show vbModal
End Sub

Private Sub CrSchMnu_Click()
frmShiftCr.Show vbModal
End Sub

Private Sub DailyPMnu_Click()
    Call DoDaily
End Sub

Private Sub DecHDayMnu_Click()
    frmDecH.Show vbModal
End Sub

Private Sub DepMastMnu_Click()
    frmDept.Show vbModal
End Sub

Private Sub EmpMastMnu_Click()
    frmEmp.Show vbModal
End Sub

Private Sub EncLvMnu_Click()
bytFormToLoad = 3       '' Encash Leaves
PassFrm.Show vbModal
End Sub

Private Sub ExtMnu_Click()
    Unload Me
End Sub

Private Sub GetPvstarVars()
On Error GoTo ERR_P
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "select lvcode,custcode,leave from leavdesc where leave = 'Present Days' " & _
"or leave = 'Absent Days' " & _
"or leave = 'Weekly Off' Or leave = 'Holiday Days' order by leave" _
, ConMain, adOpenDynamic, adLockOptimistic
If adRsInstall("definod") = False Then
    With pVStar
        .AbsCode = adrsLeave("lvcode")
        adrsLeave.MoveNext
        .HlsCode = adrsLeave("lvcode")
        adrsLeave.MoveNext
        .PrsCode = adrsLeave("lvcode")
        adrsLeave.MoveNext
        .WosCode = adrsLeave("lvcode")
    End With
Else
    adrsLeave.MoveFirst
    Do While Not adrsLeave.EOF
        Select Case UCase(Trim(adrsLeave("leave")))
            Case "PRESENT DAYS": pVStar.PrsCode = adrsLeave("custcode")
            Case "ABSENT DAYS": pVStar.AbsCode = adrsLeave("custcode")
            Case "WEEKLY OFF": pVStar.WosCode = adrsLeave("custcode")
            Case "HOLIDAY DAYS": pVStar.HlsCode = adrsLeave("custcode")
            Case Else
        End Select
        adrsLeave.MoveNext
    Loop
End If
typVar.strAbsCode = pVStar.AbsCode
typVar.strPrsCode = pVStar.PrsCode
typVar.strHlsCode = pVStar.HlsCode
typVar.strWosCode = pVStar.WosCode
Exit Sub
ERR_P:
    ShowError ("Pvstar Statuses Failed")
    Resume Next
End Sub

Public Sub DoDaily()
On Error GoTo ERR_P
If adrsRits.State = 1 Then adrsRits.Close
adrsRits.Open "Select * from Exc", ConMain, adOpenKeyset, adLockOptimistic
If adrsRits("Daily") = 1 Then
    adrsRits.Close
    MsgBox NewCaptionTxt("53047", adrsC), vbExclamation
Else
    adrsRits("Daily") = 1
    adrsRits.Update
    adrsRits.Close
    frmDailyTry.Show vbModal
    Call SetDailyFlag        '' Reset Flag
End If
Exit Sub
ERR_P:
    ShowError ("Daily Process :: " & Me.Caption)
    'Resume Next
End Sub

