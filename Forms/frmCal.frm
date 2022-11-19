VERSION 5.00
Begin VB.Form frmCal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   90
      ScaleHeight     =   2325
      ScaleWidth      =   4005
      TabIndex        =   0
      Top             =   390
      Width           =   4005
      Begin VB.Shape shpWD 
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   4005
      End
      Begin VB.Shape Shape1 
         Height          =   1980
         Left            =   0
         Top             =   330
         Width           =   4005
      End
      Begin VB.Line Line1 
         X1              =   570
         X2              =   570
         Y1              =   0
         Y2              =   2300
      End
      Begin VB.Line Line2 
         X1              =   1140
         X2              =   1140
         Y1              =   0
         Y2              =   2300
      End
      Begin VB.Line Line3 
         X1              =   1710
         X2              =   1710
         Y1              =   0
         Y2              =   2300
      End
      Begin VB.Line Line4 
         X1              =   2280
         X2              =   2280
         Y1              =   0
         Y2              =   2300
      End
      Begin VB.Line Line5 
         X1              =   2850
         X2              =   2850
         Y1              =   0
         Y2              =   2300
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   54
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblTue 
         Alignment       =   2  'Center
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   53
         Top             =   30
         Width           =   525
      End
      Begin VB.Label lblWed 
         Alignment       =   2  'Center
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   52
         Top             =   30
         Width           =   525
      End
      Begin VB.Label lblThu 
         Alignment       =   2  'Center
         Caption         =   "Thu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1740
         TabIndex        =   51
         Top             =   30
         Width           =   525
      End
      Begin VB.Label lblFri 
         Alignment       =   2  'Center
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2310
         TabIndex        =   50
         Top             =   30
         Width           =   525
      End
      Begin VB.Label lblSat 
         Alignment       =   2  'Center
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   30
         Width           =   525
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   4000
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   48
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   2
         Left            =   1170
         TabIndex        =   47
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   3
         Left            =   1740
         TabIndex        =   46
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   4
         Left            =   2310
         TabIndex        =   45
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   5
         Left            =   2880
         TabIndex        =   44
         Top             =   360
         Width           =   525
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   4000
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   6
         Left            =   3450
         TabIndex        =   43
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   7
         Left            =   30
         TabIndex        =   42
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   8
         Left            =   600
         TabIndex        =   41
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   9
         Left            =   1170
         TabIndex        =   40
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   10
         Left            =   1740
         TabIndex        =   39
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   11
         Left            =   2310
         TabIndex        =   38
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   12
         Left            =   2880
         TabIndex        =   37
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   13
         Left            =   3450
         TabIndex        =   36
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   14
         Left            =   30
         TabIndex        =   35
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   15
         Left            =   600
         TabIndex        =   34
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   16
         Left            =   1170
         TabIndex        =   33
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   17
         Left            =   1740
         TabIndex        =   32
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   18
         Left            =   2310
         TabIndex        =   31
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   19
         Left            =   2880
         TabIndex        =   30
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   20
         Left            =   3450
         TabIndex        =   29
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   21
         Left            =   30
         TabIndex        =   28
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   22
         Left            =   600
         TabIndex        =   27
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   23
         Left            =   1170
         TabIndex        =   26
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   24
         Left            =   1740
         TabIndex        =   25
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   25
         Left            =   2310
         TabIndex        =   24
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   26
         Left            =   2880
         TabIndex        =   23
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   27
         Left            =   3450
         TabIndex        =   22
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   28
         Left            =   30
         TabIndex        =   21
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   29
         Left            =   600
         TabIndex        =   20
         Top             =   1680
         Width           =   525
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   4000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   4000
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   34
         Left            =   3450
         TabIndex        =   18
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   33
         Left            =   2880
         TabIndex        =   17
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   32
         Left            =   2310
         TabIndex        =   16
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   31
         Left            =   1740
         TabIndex        =   15
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   30
         Left            =   1170
         TabIndex        =   14
         Top             =   1680
         Width           =   525
      End
      Begin VB.Line Line12 
         X1              =   3420
         X2              =   3420
         Y1              =   0
         Y2              =   2300
      End
      Begin VB.Label lblSun 
         Alignment       =   2  'Center
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3450
         TabIndex        =   13
         Top             =   30
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   35
         Left            =   30
         TabIndex        =   12
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   36
         Left            =   600
         TabIndex        =   11
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   37
         Left            =   1170
         TabIndex        =   10
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   38
         Left            =   1740
         TabIndex        =   9
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   39
         Left            =   2310
         TabIndex        =   8
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   40
         Left            =   2880
         TabIndex        =   7
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   285
         Index           =   41
         Left            =   3450
         TabIndex        =   6
         Top             =   2010
         Width           =   525
      End
      Begin VB.Line Line11 
         X1              =   0
         X2              =   4000
         Y1              =   1980
         Y2              =   1980
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2940
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   1710
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   825
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label lblMonYea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEPTEMBER 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   1950
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' Record Variables
Dim intCurIndex As Integer
Dim bytCalMode As Byte
Dim bytFirstIndex As Byte
Dim bytLastIndex As Byte
Dim bytMonthEnd As Byte
'' Constants
Const MAX_YEAR = 2199
Const MIN_YEAR = 1901
''
Dim adrsC As New ADODB.Recordset

Private Sub cboMonth_Click()
On Error Resume Next
If bytCalMode = 0 Then Exit Sub
Call FillCalendar(0, cboMonth.ListIndex + 1, CInt(cboYear.Text))
Call SetLabel(cboMonth.Text, cboYear.Text)
End Sub

Private Sub cboYear_Click()
On Error Resume Next
If bytCalMode = 0 Then Exit Sub
Call FillCalendar(0, cboMonth.ListIndex + 1, CInt(cboYear.Text))
Call SetLabel(cboMonth.Text, cboYear.Text)
End Sub

Private Sub cmdCan_Click()
On Error Resume Next
Unload Me
Screen.ActiveControl = varCalDt
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
If intCurIndex = -1 Then
    MsgBox NewCaptionTxt("09002", adrsC), vbExclamation
    Exit Sub
End If
Call DateCopy
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
On Error GoTo LBL
Call SetFormIcon(Me, True)            '' Sets the Forms Icon
Call RetCaptions                 '' Sets the Forms Captions
Dim intTmp As Integer
'' Load Mode
bytCalMode = 0
'' Fill the Month
For intTmp = 1 To 12
    cboMonth.AddItem Choose(intTmp, "January", "February", "March", "April", "May", "June", _
    "July", "August", "September", "October", "November", "December")
Next
'' End
'' Fill the Year
For intTmp = MIN_YEAR To MAX_YEAR
    cboYear.AddItem intTmp
Next
intCurIndex = -1
If GotValidDate Then
    cboMonth.ListIndex = Month(varCalDt) - 1
    cboYear.Text = CStr(Year(varCalDt))
    Call FillCalendar(Day(varCalDt), cboMonth.ListIndex + 1, Val(cboYear.Text))
    Call SetLabel(cboMonth.Text, cboYear.Text)
Else
LBL:
    cboMonth.ListIndex = Month(Date) - 1
    intTmp = Year(Date)
    cboYear.Text = intTmp
    Call FillCalendar(Day(Date), Month(Date), intTmp)
    Call SetLabel(cboMonth.Text, cboYear.Text)
End If
'' End
'' End Load Mode
bytCalMode = 1
End Sub

Private Sub ChangeState(ByRef Obj As Object, Optional bytChecked As Byte = 0)
On Error Resume Next
Select Case bytChecked
    Case 0          '' UnSelected
        Obj.Enabled = True
        Obj.BorderStyle = 0
        Obj.BackColor = &H8000000F
        Obj.ForeColor = &H80000012
    Case 1          '' Selected
        Obj.Enabled = True
        Obj.BorderStyle = 1
        Obj.BackColor = &H8080&
        Obj.ForeColor = &H80FFFF
    Case 2          '' Not in Current Month
        Obj.Enabled = True
        Obj.BorderStyle = 0
        Obj.BackColor = &H8000000F
        Obj.ForeColor = &H8000000C
    Case 3
        Obj.Enabled = False
        Obj.ForeColor = vbWhite
    Case 4
        Obj.Enabled = True
End Select
End Sub

Private Function GotValidDate() As Boolean
On Error GoTo ERR_P
Dim bytD As Byte, bytM As Byte, intY As Integer
GotValidDate = False
'' On Length
Select Case Len(varCalDt)
    Case 8, 10
        '' Check for Valida date
        bytD = Day(varCalDt)
        bytM = Month(varCalDt)
        intY = Year(varCalDt)
        If intY > MAX_YEAR Then Exit Function
        If intY < MIN_YEAR Then Exit Function
        If bytM <= 0 And bytM > 12 Then Exit Function
        If bytD <= 0 Then Exit Function
        Select Case bytM
            Case 1, 3, 5, 7, 8, 10, 12
                If bytD > 31 Then Exit Function
            Case 2
                If LeapOrNot(intY) Then
                    If bytD > 29 Then Exit Function
                Else
                    If bytD > 28 Then Exit Function
                End If
            Case 4, 6, 9, 11
                If bytD > 30 Then Exit Function
        End Select
        GotValidDate = True
End Select
Exit Function
ERR_P:
    GotValidDate = False
End Function

Private Sub SetLabel(ByVal strMonth As String, intYear As Integer)
On Error Resume Next
lblMonYea.Caption = strMonth & "   " & intYear
End Sub

Private Sub FillCalendar(ByVal bytDay As Byte, ByVal bytMonth As Byte, ByVal intYear As Integer)
On Error Resume Next
Dim bytTmp As Byte
'' Get the Index of the First day
Select Case UCase(Left(WeekdayName(WeekDay(Year_Start(bytMonth, intYear), vbUseSystemDayOfWeek)), 3))
    Case "MON"
        bytFirstIndex = 0
    Case "TUE"
        bytFirstIndex = 1
    Case "WED"
        bytFirstIndex = 2
    Case "THU"
        bytFirstIndex = 3
    Case "FRI"
        bytFirstIndex = 4
    Case "SAT"
        bytFirstIndex = 5
    Case "SUN"
        bytFirstIndex = 6
End Select
Call GetMonthEnd(bytMonth, intYear)
bytLastIndex = bytFirstIndex + bytMonthEnd - 1
'' Fill the Calendar for the Given Month
For bytTmp = bytFirstIndex To bytLastIndex
    '' Current Day Index
    If bytDay = bytTmp - (bytFirstIndex - 1) Then intCurIndex = bytTmp
    '' Setting the Caption
    Label1(bytTmp).Caption = bytTmp - (bytFirstIndex - 1)
    Call ChangeState(Label1(bytTmp))
Next
'' Fill the Previous Recs
Call PrevRecs(bytMonth, intYear)
Call NextRecs(bytMonth, intYear)
If bytDay = 0 Then intCurIndex = -1
If intCurIndex <> -1 Then
    Call ChangeState(Label1(intCurIndex), 1)
End If
End Sub

Private Sub GetMonthEnd(ByVal bytMonth As Byte, intYear As Integer)
On Error Resume Next
Select Case bytMonth
    Case 1, 3, 5, 7, 8, 10, 12
        bytMonthEnd = 31
    Case 2
        bytMonthEnd = IIf(LeapOrNot(intYear), 29, 28)
    Case 4, 6, 9, 11
        bytMonthEnd = 30
End Select
End Sub

Private Function LeapOrNot(ByVal intYear As Integer) As Boolean   '' Leap year Checking
On Error Resume Next
If intYear Mod 4 = 0 Then                   '' If Divisible by 4
    ' Is it a Century?
    If intYear Mod 100 = 0 Then             '' if Divisible by 100
        ' If a Century, must be Evenly Divisible by 400.
        If intYear Mod 400 = 0 Then         '' If Divisible by 400
            LeapOrNot = True                '' Leap Year
        Else
            LeapOrNot = False               '' Non-Leap Year
        End If
    Else
        LeapOrNot = True                    '' Leap Year
    End If
Else
    LeapOrNot = False                       '' Non-Leap Year
End If
End Function

Private Sub PrevRecs(ByVal bytMonth As Byte, ByVal intYear As Integer)
On Error Resume Next
Dim intTmp As Integer
'' Settle Previous Records
If bytFirstIndex > 0 Then
    If bytMonth = 1 Then
        bytMonth = 12
        intYear = intYear - 1
    Else
        bytMonth = bytMonth - 1
    End If
    Call GetMonthEnd(bytMonth, intYear)
    '' Code for Caption
    For intTmp = bytFirstIndex - 1 To 0 Step -1
        Label1(intTmp).Caption = bytMonthEnd
        Call ChangeState(Label1(intTmp), 2)
        If intYear < MIN_YEAR Then
            Call ChangeState(Label1(intTmp), 3)
        Else
            Call ChangeState(Label1(intTmp), 4)
        End If
        bytMonthEnd = bytMonthEnd - 1
    Next
End If
End Sub

Private Sub NextRecs(ByVal bytMonth As Byte, ByVal intYear As Integer)
On Error Resume Next
Dim bytTmp As Byte
'' Settle Previous Records
If bytLastIndex < Label1.Count - 2 Then
    If bytMonth = 12 Then
        bytMonth = 1
        intYear = intYear + 1
    Else
        bytMonth = bytMonth + 1
    End If
    Call GetMonthEnd(bytMonth, intYear)
    '' Code for Caption
    For bytTmp = bytLastIndex + 1 To Label1.Count - 1
        Label1(bytTmp).Caption = bytTmp - (bytLastIndex)
        Call ChangeState(Label1(bytTmp), 2)
        If intYear > MAX_YEAR Then
            Call ChangeState(Label1(bytTmp), 3)
        Else
            Call ChangeState(Label1(bytTmp), 4)
        End If
    Next
End If
End Sub

Private Sub Label1_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case Is < bytFirstIndex
        If cboMonth.ListIndex = 0 Then
            Call FillCalendar(Val(Label1(Index).Caption), 12, Val(cboYear.Text) - 1)
            Call SetListMonth(11)
            Call SetListYear(cboYear.ListIndex - 1)
        Else
            Call FillCalendar(Val(Label1(Index).Caption), cboMonth.ListIndex, Val(cboYear.Text))
            Call SetListMonth(cboMonth.ListIndex - 1)
        End If
    Case bytFirstIndex To bytLastIndex
        If intCurIndex <> -1 Then Call ChangeState(Label1(intCurIndex), 0)
        Call ChangeState(Label1(Index), 1): intCurIndex = Index
    Case Is > bytLastIndex
        If cboMonth.ListIndex = 11 Then
            Call FillCalendar(Val(Label1(Index).Caption), 1, Val(cboYear.Text) + 1)
            Call SetListMonth(0)
            Call SetListYear(cboYear.ListIndex + 1)
        Else
            Call FillCalendar(Val(Label1(Index).Caption), cboMonth.ListIndex + 2, Val(cboYear.Text))
            Call SetListMonth(cboMonth.ListIndex + 1)
        End If
End Select
Call SetLabel(cboMonth.Text, cboYear.Text)
Picture1.SetFocus
End Sub

Private Sub Label1_DblClick(Index As Integer)
Call cmdOK_Click
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If intCurIndex = -1 Then Exit Sub
Select Case KeyCode
    Case vbKeyLeft
        Call KeyLeft
    Case vbKeyRight
        Call KeyRight
    Case vbKeyUp
        Call KeyUp
    Case vbKeyDown
        Call KeyDown
End Select
End Sub

Private Sub KeyLeft()
On Error Resume Next
Select Case intCurIndex
    Case 0                                              '' Case A [IntCurIndex =  0]
        If cboMonth.ListIndex = 0 Then                  '' Case 01. Month =  January
            If cboYear.ListIndex = 0 Then Exit Sub      '' If Minimum Date
            Call FillCalendar(31, 12, Val(cboYear.Text) - 1)
            Call SetListMonth(11)
            Call SetListYear(cboYear.ListIndex - 1)
        Else                                            '' Case 02. Month <> January
            Call GetMonthEnd(cboMonth.ListIndex, Val(cboYear.Text))
            Call FillCalendar(bytMonthEnd, cboMonth.ListIndex, Val(cboYear.Text))
            Call SetListMonth(cboMonth.ListIndex - 1)
        End If
    Case Else                                           '' Case B [IntCurIndex <> 0]
        If cboMonth.ListIndex = 0 Then                  '' Case 01. Month =  January
            If Val(Label1(intCurIndex - 1).Caption) > _
            Val(Label1(intCurIndex).Caption) Then       '' If the Caption is Greater
                If Val(cboYear.Text) = MIN_YEAR Then Exit Sub
                Call FillCalendar(31, 12, Val(cboYear.Text) - 1)
                Call SetListMonth(11)
                Call SetListYear(cboYear.ListIndex - 1)
            Else
                Call ChangeState(Label1(intCurIndex), 0)
                intCurIndex = intCurIndex - 1
                Call ChangeState(Label1(intCurIndex), 1)
            End If
        Else                                            '' Case 02. Month <> January
            If Val(Label1(intCurIndex - 1).Caption) > _
            Val(Label1(intCurIndex).Caption) Then       '' If the Caption is Greater
                Call GetMonthEnd(cboMonth.ListIndex, Val(cboYear.Text))
                Call FillCalendar(bytMonthEnd, cboMonth.ListIndex, Val(cboYear.Text))
                Call SetListMonth(cboMonth.ListIndex - 1)
            Else
                Call ChangeState(Label1(intCurIndex), 0)
                intCurIndex = intCurIndex - 1
                Call ChangeState(Label1(intCurIndex), 1)
            End If
        End If
End Select
Call SetLabel(cboMonth.Text, cboYear.Text)
End Sub

Private Sub SetListMonth(ByVal bytIndex As Byte)
On Error Resume Next
bytCalMode = 0
cboMonth.ListIndex = bytIndex
bytCalMode = 1
End Sub

Private Sub SetListYear(ByVal intIndex As Integer)
On Error Resume Next
bytCalMode = 0
cboYear.ListIndex = intIndex
bytCalMode = 1
End Sub

Private Sub KeyRight()
On Error Resume Next
Select Case intCurIndex
    Case Is <= 36                                        '' Case B [IntCurIndex <= 36]
        If cboYear.Text = CStr(MAX_YEAR) And _
        Label1(intCurIndex).Caption = "31" Then Exit Sub '' If Maximum Date
        If cboMonth.ListIndex = 11 Then                  '' Case 01. Month =  December
            If Val(Label1(intCurIndex + 1).Caption) < _
            Val(Label1(intCurIndex).Caption) Then       '' If the Caption is Lesser
                Call FillCalendar(1, 1, Val(cboYear.Text) + 1)
                Call SetListMonth(0)
                Call SetListYear(cboYear.ListIndex + 1)
            Else
                Call ChangeState(Label1(intCurIndex), 0)
                intCurIndex = intCurIndex + 1
                Call ChangeState(Label1(intCurIndex), 1)
            End If
        Else                                            '' Case 02. Month <> December
            If Val(Label1(intCurIndex + 1).Caption) < _
            Val(Label1(intCurIndex).Caption) Then       '' If the Caption is Lesser
                Call GetMonthEnd(cboMonth.ListIndex + 2, Val(cboYear.Text))
                Call FillCalendar(1, cboMonth.ListIndex + 2, Val(cboYear.Text))
                Call SetListMonth(cboMonth.ListIndex + 1)
            Else
                Call ChangeState(Label1(intCurIndex), 0)
                intCurIndex = intCurIndex + 1
                Call ChangeState(Label1(intCurIndex), 1)
            End If
        End If
End Select
Call SetLabel(cboMonth.Text, cboYear.Text)
End Sub

Private Sub KeyUp()
On Error Resume Next
Dim bytTmp As Byte
bytTmp = Val(Label1(intCurIndex).Caption)
Select Case bytTmp
    Case Is > 7
        Call ChangeState(Label1(intCurIndex), 0)
        intCurIndex = intCurIndex - 7
        Call ChangeState(Label1(intCurIndex), 1)
    Case Else
        If cboMonth.ListIndex = 0 Then
            If cboYear.ListIndex = 0 Then Exit Sub
            Call FillCalendar(24 + bytTmp, 12, Val(cboYear.Text) - 1)
            Call SetListMonth(11)
            Call SetListYear(cboYear.ListIndex - 1)
        Else
            Call GetMonthEnd(cboMonth.ListIndex, Val(cboYear.Text))
            Call FillCalendar(bytMonthEnd - (7 - bytTmp), cboMonth.ListIndex, Val(cboYear.Text))
            Call SetListMonth(cboMonth.ListIndex - 1)
        End If
End Select
Call SetLabel(cboMonth.Text, cboYear.Text)
''  >7 change status -7
''  1 to 7
''      A. =  January
''          01. If 1900 Exit sub
''          02. (GetMonthEnd-(7 - Caption)),Month 12, Year -1
''      B. <> January
''          02. (GetMonthEnd-(7 - Caption)),Month -1, Year
End Sub

Private Sub KeyDown()
On Error Resume Next
Dim bytTmp As Byte
Call GetMonthEnd(cboMonth.ListIndex + 1, Val(cboYear.Text))
bytTmp = Val(Label1(intCurIndex).Caption)
Select Case bytTmp
    Case Is <= (bytMonthEnd - 7)
        Call ChangeState(Label1(intCurIndex), 0)
        intCurIndex = intCurIndex + 7
        Call ChangeState(Label1(intCurIndex), 1)
    Case Else
        If cboMonth.ListIndex = 11 Then
            If cboYear.ListIndex = cboYear.ListCount - 1 Then Exit Sub
            Call FillCalendar(bytTmp - 24, 1, Val(cboYear.Text) + 1)
            Call SetListMonth(0)
            Call SetListYear(cboYear.ListIndex + 1)
        Else
            Call FillCalendar(7 - CInt(bytMonthEnd - bytTmp), cboMonth.ListIndex + 2, Val(cboYear.Text))
            Call SetListMonth(cboMonth.ListIndex + 1)
        End If
End Select
Call SetLabel(cboMonth.Text, cboYear.Text)
End Sub

Private Sub DateCopy()
On Error GoTo ERR_P
Dim DateText As Date
DateText = Label1(intCurIndex).Caption & "/" & cboMonth.Text & "/" & cboYear.Text
If bytDateF = 1 Then
    Clipboard.SetText Format(DateText, "MM/DD/YYYY"), vbCFText
Else
    Clipboard.SetText DateDisp(DateText), vbCFText
End If
Unload Me
If TypeOf Screen.ActiveControl Is TextBox Then
    Screen.ActiveControl = Clipboard.GetText(vbCFText)
End If
DateText = 0
Exit Sub
ERR_P:
    ShowError ("Date Copy :: " & Me.Caption)
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '09%'", ConMain, adOpenStatic
Me.Caption = NewCaptionTxt("09001", adrsC)
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)
cmdCan.Caption = NewCaptionTxt("00003", adrsMod)
End Sub

