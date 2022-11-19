VERSION 5.00
Begin VB.Form frmLvBalOption 
   Caption         =   "Leave Balance Options"
   ClientHeight    =   1560
   ClientLeft      =   5490
   ClientTop       =   4455
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   3240
   Begin VB.Frame fraLvBalOption 
      Caption         =   "Select Leave  Balance Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   3375
      Begin VB.Frame fraDt 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   3135
         Begin VB.TextBox txtDaily 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "D"
            Top             =   90
            Width           =   1125
         End
         Begin VB.Label lblDaily 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Report till the Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   6
            Top             =   180
            Width           =   1785
         End
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optMnthWise 
         Caption         =   "Month-Wise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optDtWise 
         Caption         =   "Date-Wise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLvBalOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub cmdOK_Click()
Dim SelMnth As String, SelYear As String
Dim PrvMnth As String, PrvYear As String
Dim Mnth As String, dt As String, Yr As String
Dim Ldate As String, Fdate As String
Dim strMon_Trn As String
Dim strFileName As String

If optDtWise.Value = True Then
    typDlyLvBal.DailyDt = txtDaily.Text
End If

If typDlyLvBal.bytDtOpt = 1 Then
    dt = Day(typDlyLvBal.DailyDt)
    Mnth = Month(typDlyLvBal.DailyDt)
    Yr = Year(typDlyLvBal.DailyDt)
    typDlyLvBal.strMnth = MonthName(Mnth)
    typDlyLvBal.strYr = Yr
    Ldate = DateCompStr(FdtLdt(MonthNumber(typDlyLvBal.strMnth), typDlyLvBal.strYr, "L"))
    If typDlyLvBal.bytDtOpt = 1 Then
         If dt = Day(Ldate) Then
            Fdate = DateCompStr(FdtLdt(MonthNumber(typDlyLvBal.strMnth), typDlyLvBal.strYr, "F"))
            Ldate = DateCompStr(FdtLdt(MonthNumber(typDlyLvBal.strMnth), typDlyLvBal.strYr, "L"))
        Else
            If Mnth = 1 Then
                Mnth = 12
                Yr = Yr - 1
                Fdate = DateCompStr(FdtLdt(MonthNumber(MonthName(Mnth)), Yr, "F"))
                Ldate = DateCompStr(FdtLdt(MonthNumber(MonthName(Mnth)), Yr, "L"))
            Else
                Mnth = Mnth - 1
                Fdate = DateCompStr(FdtLdt(MonthNumber(MonthName(Mnth)), Yr, "F"))
                Ldate = DateCompStr(FdtLdt(MonthNumber(MonthName(Mnth)), Yr, "L"))
            End If
        End If
    End If

    If Val(pVStar.Yearstart) > MonthNumber(typDlyLvBal.strMnth) Then
        strFileName = "lvinfo" & Right(CStr(CInt(typDlyLvBal.strYr) - 1), 2)
    Else
        strFileName = "lvinfo" & Right(typDlyLvBal.strYr, 2)
    End If
    
    If Not FindTable(strFileName) Then MsgBox NewCaptionTxt("40082", adrsC), vbInformation
    
    strMon_Trn = MakeName(MonthName(Month(DateCompDate(txtDaily.Text))), _
    Year(DateCompDate(txtDaily.Text)), "trn")
    If Not FindTable(strMon_Trn) Then
    'Call SetMSF1Cap(10)
        MsgBox NewCaptionTxt("40065", adrsC) & _
        MonthName(Month(DateCompDate(txtDaily.Text))), vbExclamation
        txtDaily.SetFocus
    End If

    typDlyLvBal.typFdate = Fdate
    typDlyLvBal.typLdate = Ldate
End If
Me.Hide
End Sub

Private Sub Form_Load()
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '40%'", ConMain, adOpenStatic, adLockReadOnly
'Call SetFormIcon(Me, True)        '' Sets the Forms Icon
fraDt.Visible = False
optMnthWise.Value = True
txtDaily.Text = DateDisp(Date)
optMnthWise.Value = True
End Sub

Private Sub optDtWise_Click()
fraDt.Visible = True
typDlyLvBal.bytMnthOpt = 0
typDlyLvBal.bytDtOpt = 1
End Sub

Private Sub optMnthWise_Click()
fraDt.Visible = False
typDlyLvBal.bytMnthOpt = 1
typDlyLvBal.bytDtOpt = 0
End Sub


Private Sub txtDaily_Click()
varCalDt = ""
varCalDt = Trim(txtDaily.Text)
txtDaily.Text = ""
Call ShowCalendar
If Not dlyValid Then Call SetMSF1Cap(1): Exit Sub
cmdOk.SetFocus
End Sub
Private Function dlyValid() As Boolean
On Error GoTo ERR_P
dlyValid = True                                 '' FUNCTION FOR DAILY REPORT VALIDATIONS
Dim strFileName As String
'strFileName = Left(MonthName(Month(txtDaily.Text)), 3) & Right(CStr(CInt(Year(txtDaily.Text)) - 1), 2) & "trn"
strFileName = Left(MonthName(Month(txtDaily.Text)), 3) & Right(Year(txtDaily.Text), 2) & "trn"

If Not FindTable(strFileName) Then
    'Call SetMSF1Cap(10)
    MsgBox NewCaptionTxt("40065", adrsC) & _
        MonthName(Month(DateCompDate(txtDaily.Text))), vbExclamation
    txtDaily.SetFocus
    dlyValid = False
    Exit Function
End If
Exit Function
ERR_P:
    ShowError ("dlyValid :: Reportsfrm")
    dlyValid = False
End Function

Private Sub txtDaily_Validate(Cancel As Boolean)
If Not ValidDate(txtDaily) Then
    txtDaily.SetFocus
    Cancel = True
End If
End Sub
Private Sub txtDaily_GotFocus()
    Call GF(txtDaily)
End Sub

Private Sub txtDaily_KeyPress(KeyAscii As Integer)
    Call CDK(txtDaily, KeyAscii)
End Sub

