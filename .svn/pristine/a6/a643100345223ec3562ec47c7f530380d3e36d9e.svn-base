VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPaySlip 
   Caption         =   "Salary Slip"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Salary Slip"
   LockControls    =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6195
      Begin VB.ComboBox cmbMonth 
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
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1425
      End
      Begin MSForms.ComboBox cboTo 
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   960
         Width           =   1125
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1984;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFrom 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   1125
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1984;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3120
         TabIndex        =   9
         Top             =   960
         Width           =   210
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report for the month of"
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
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   2205
      End
   End
   Begin VB.Frame Frame 
      Height          =   855
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPaySlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo Err
    Dim crxApp As CRAXDRT.Application
    Dim crxRpt As CRAXDRT.Report
    
    Set crxApp = New CRAXDRT.Application
    Set crxRpt = crxApp.OpenReport(App.path & "\Reports\mstsalaryslip.rpt", 1)
    Dim sql As String
    Dim strlvfile As String
    If cmbMonth.Text = "" Then Exit Sub
    strlvfile = "lvtrn" & Right(pVStar.YearSel, 2)
    
    sql = "SELECT Empmst.empcode, Empmst.name, frmDesignation.DesigName, deptdesc.desc, Empmst.salary, " & strlvfile & ".[A ], " & strlvfile & ".HL, " & strlvfile & ".[P ], " & strlvfile & ".WO, " & strlvfile & ".paiddays, " & strlvfile & ".lst_date, company.CName"
    sql = sql + " FROM company INNER JOIN ((deptdesc INNER JOIN (" & strlvfile & " INNER JOIN Empmst ON " & strlvfile & ".empcode = Empmst.empcode) ON deptdesc.dept = Empmst.dept) INNER JOIN frmDesignation ON Empmst.designatn = frmDesignation.DesigCode) ON company.Company = Empmst.company "
    sql = sql + " Where  Month(lst_date) =  " & cmbMonth.ListIndex + 1
    If cboFrom.Text <> "" Or cboTo.Text <> "" Then sql = sql + " And Empmst.Empcode >= '" & cboFrom.Text & "' and Empmst.Empcode <= '" & cboTo.Text & "'"
    sql = sql + " Order by Empmst.Empcode"
    If adrsCrep.State = 1 Then adrsCrep.Close
    adrsCrep.Open sql, ConMain, adOpenStatic, adLockOptimistic
    
    crxRpt.Database.SetDataSource adrsCrep
    frmCRV.Caption = "Employee List Report"
    'Set Report = crxApp.OpenReport(App.path & "\Reports\dlyArrival.rpt", 1)
    
    Set CRV = frmCRV.CRV:  CRV.ReportSource = crxRpt
    CRV.ViewReport
    frmCRV.Show vbModal
    Exit Sub
Err:
        ShowError ("Pay Slip Report Priview :: " & Me.Caption)
End Sub

Private Sub Form_Load()
With cmbMonth           '' Month
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
    FillCombos
End Sub

Private Sub FillCombos()
    Call ComboFill(cboFrom, 19, 2)
    Call ComboFill(cboTo, 19, 2)
    cboTo.ListIndex = cboTo.ListCount - 1
    cboFrom.ListIndex = 0
    Exit Sub
ERR_P:
    ShowError ("FillCombos::" & Me.Caption)
End Sub

