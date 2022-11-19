VERSION 5.00
Begin VB.Form MarkFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mark All Employee as present"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ExitCmd 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton OKcmd 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox LastDateTxt 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   " "
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Datetxt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Tag             =   "D"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "to the end of month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mark all employee as present from the date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   3540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last done for the date "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2> Use this option once for particular month && year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Width           =   4140
   End
   Begin VB.Label Label1 
      Caption         =   "Note : 1> First make sure you have finished monthly processing."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5370
   End
End
Attribute VB_Name = "MarkFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LDate As String, FDate As String, FileName As String

Private Sub Datetxt_Click()
varCalDt = ""
varCalDt = Trim(Datetxt.Text)
Datetxt.Text = ""
Load CalendarFrm
CalendarFrm.Show 1
End Sub

Private Sub Datetxt_GotFocus()
    Call GF(Datetxt)
End Sub

Private Sub Datetxt_KeyPress(KeyAscii As Integer)
    Call CDK(Datetxt, KeyAscii)
End Sub

Private Sub Datetxt_Validate(Cancel As Boolean)
    If Not ValidDate(Datetxt) Then Datetxt.SetFocus: Cancel = True
End Sub

Private Sub ExitCmd_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo ERR_P
If UCase(Trim(userName)) <> strPrintUser Then
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "select lv_rights from user_leave_rights where username=" & "'" & userName & "'" _
    , VstarDataEnv.cnDJConn
    If Not (adrsTemp.BOF And adrsTemp.EOF) Then
        If Mid(adrsTemp(0), 26, 1) = "0" Then
            MsgBox "User " & userName & " does not have Rights to Perform  this Operation.", vbInformation, App.EXEName
            Unload Me
        End If
    Else
        MsgBox "Invalid User", vbExclamation, App.EXEName & "  :: User Error "
        Unload Me
    End If
    adrsTemp.Close
End If
Datetxt.Enabled = True
Datetxt.SetFocus
Exit Sub
ERR_P:
    ShowError ("Activate :: " & Me.Caption)
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)        '' Sets the Forms Icon
Datetxt.Enabled = False
Call SetToolTipText(Me)     '' Set the ToolTipText
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open "select Lastdate from Markpres", _
VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic
If Not (adrsTemp.EOF And adrsTemp.BOF) Then
    If Not IsNull(adrsTemp("LastDate")) Then LastDateTxt = DateDisp(adrsTemp!LastDate)
End If
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub Okcmd_Click()
On Error GoTo ERR_Particular
Dim DayDiff%, dt As Date
If Trim(Datetxt) = "" Then
    MsgBox "Please mention Date", vbExclamation, App.EXEName
    Datetxt.SetFocus
    Exit Sub
End If
dt = DateCompDate(Datetxt.Text)
LDate = FdtLdt(Month(dt), CStr(Year(dt)), "l")
DayDiff = DateDiff("d", dt, DateCompDate(LDate)) + 1
FileName = Left(MonthName(Month(dt)), 3) & Right(Year(dt), 2) & "trn"
If FindTable(FileName) Then
    FileName = "lvtrn" & Right(strYearFrom(LDate), 2)
    If FindTable(FileName) Then
        If MsgBox("You want to make all employee(s) present for " & DayDiff & " Days", _
        vbQuestion + vbYesNo, App.EXEName) = vbYes Then
            Call AddActivityLog(lg_NoModeAction, 3, 27)     '' Mark Present Log
            VstarDataEnv.cnDJConn.Execute "Update " & FileName & " Set Paiddays=PaidDays+" & _
            DayDiff & " Where Lst_Date=" & strDTEnc & DateCompStr(LDate) & strDTEnc & _
            " and EmpCode" & " in (Select EmpMst.EmpCode From EmpMst," & FileName & _
            " Where EmpMst.EmpCode=" & FileName & ".EmpCode and (EmpMst.Leavdate>" & _
            strDTEnc & DateCompStr(LDate) & strDTEnc & " or EmpMst.Leavdate is Null))"
            '' Update Last Date
            VstarDataEnv.cnDJConn.Execute "Update Markpres set Lastdate=" & strDTEnc & _
            DateSaveIns(dt) & strDTEnc
            MsgBox "All Employees Marked Present for " & DayDiff & " Days", _
            vbExclamation, App.EXEName
        End If
    Else
        MsgBox "Yearly Leave Transaction not done for the Month of " & MonthName(Month(FDate)) & " " & Year(FDate) & vbCrLf & _
        "Cannot proceed.", vbInformation, App.EXEName
        Exit Sub
    End If
Else
    MsgBox "Monthly transaction file for the Month " & MonthName(Month(dt)) & " does not exists." & vbCrLf & _
    "Cannot proceed", vbInformation, App.EXEName
End If
Exit Sub
ERR_Particular:
    ShowError ("Procedure Error :: " & Me.Caption)
End Sub
