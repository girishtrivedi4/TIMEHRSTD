VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPhoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Photo and Signature"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Signature photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   4050
      TabIndex        =   10
      Top             =   630
      Width           =   3930
      Begin VB.CommandButton cmdSignPhoto 
         Caption         =   "B&rows"
         Height          =   330
         Left            =   180
         TabIndex        =   13
         Top             =   3465
         Width           =   1140
      End
      Begin VB.CommandButton cmdSaveSign 
         Caption         =   "S&ave"
         Height          =   330
         Left            =   1395
         TabIndex        =   12
         Top             =   3465
         Width           =   1140
      End
      Begin VB.CommandButton cmdClearSign 
         Caption         =   "C&lear"
         Height          =   330
         Left            =   2610
         TabIndex        =   11
         Top             =   3465
         Width           =   1140
      End
      Begin VB.Image PicSign 
         Height          =   3075
         Left            =   90
         Stretch         =   -1  'True
         Top             =   270
         Width           =   3750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   45
      TabIndex        =   6
      Top             =   630
      Width           =   3930
      Begin VB.CommandButton cmdEmpPhoto 
         Caption         =   "&Brows"
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   3465
         Width           =   1140
      End
      Begin VB.CommandButton cmdSaveEmp 
         Caption         =   "&Save"
         Height          =   330
         Left            =   1395
         TabIndex        =   8
         Top             =   3465
         Width           =   1140
      End
      Begin VB.CommandButton cmdClearEmp 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   2610
         TabIndex        =   7
         Top             =   3465
         Width           =   1140
      End
      Begin VB.Image picEmp 
         Height          =   3120
         Left            =   90
         Stretch         =   -1  'True
         Top             =   225
         Width           =   3750
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   6570
      TabIndex        =   4
      Top             =   135
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog ComDilog 
      Left            =   990
      Top             =   4635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSignPhoto 
      Height          =   375
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4635
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtEmpPhoto 
      Height          =   375
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4635
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   330
      Left            =   5175
      TabIndex        =   1
      Top             =   135
      Width           =   1140
   End
   Begin MSForms.ComboBox cboCode 
      Height          =   375
      Left            =   1620
      TabIndex        =   5
      Top             =   90
      Width           =   3330
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label frmFoto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   450
      TabIndex        =   0
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "frmPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrStream As New ADODB.Stream

Private Sub cboCode_Change()
    Set picEmp.Picture = Nothing
    Set PicSign.Picture = Nothing
    txtEmpPhoto.Text = ""
    txtSignPhoto.Text = ""
End Sub

Private Sub cmdClearSign_Click()
    txtSignPhoto.Text = ""
    Set PicSign.Picture = Nothing
    cmdSaveSign.Enabled = True
End Sub

Private Sub cmdClearEmp_Click()
    txtEmpPhoto.Text = ""
    Set picEmp.Picture = Nothing
    cmdSaveEmp.Enabled = True
End Sub

Private Sub cmdEmpPhoto_Click()
    On Error GoTo Err
    ComDilog.Filter = "Pictures Files (*.Jpg; *.bmp)| *.Jpg;*.bmp"
    ComDilog.ShowOpen
    If ComDilog.FileName = "" Then Exit Sub
    txtEmpPhoto.Text = ComDilog.FileName
    picEmp.Picture = LoadPicture(txtEmpPhoto.Text)
    cmdSaveEmp.Enabled = True
Err:
    If Err.Number = 481 Then
       MsgBox "Invalid Photo.....", vbExclamation, "Error Loading....."
       txtEmpPhoto.Text = ""
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim rs As New ADODB.Recordset
rs.Open "Select * From EmployeePhoto Where Empcode = '" & cboCode.Text & "'", ConMain, adOpenDynamic, adLockOptimistic
Set StrStream = New ADODB.Stream
StrStream.Type = adTypeBinary
StrStream.Open
cmdSaveEmp.Enabled = False
cmdSaveSign.Enabled = False
If rs.RecordCount > 0 Then
    If IsNull(rs.Fields("EmpPhoto").Value) And IsNull(rs.Fields("SignPhoto").Value) Then
        MsgBox "No Image Found ....."
    End If
    If Not LoadPictureFromDB(rs) Then
        MsgBox "Phot Not Loaded......", vbInformation
    End If
Else
    MsgBox "No Image Found ....."
End If
End Sub

Private Sub cmdSaveEmp_Click()
    On Error GoTo procNoPicture
    Set StrStream = New ADODB.Stream
    StrStream.Type = adTypeBinary
    StrStream.Open
    Dim rsEmp As New ADODB.Recordset
    If cboCode.Text = "" Then
        MsgBox "Select The Employee", vbExclamation, "Blank Entry ......."
        cboCode.SetFocus
        Exit Sub
    ElseIf txtEmpPhoto.Text = "" Then
        If MsgBox("Do You Wnat To Save Blank Employee Photo", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        cmdEmpPhoto.SetFocus
        Exit Sub
        End If
    End If
    rsEmp.Open "Select * From EmployeePhoto Where empcode = '" & cboCode.Text & "'", ConMain, adOpenDynamic, adLockOptimistic
    If rsEmp.RecordCount < 1 Then
        rsEmp.AddNew
        rsEmp.Fields!Empcode = cboCode.Text
    End If
    
    If txtEmpPhoto.Text <> "" Then StrStream.LoadFromFile txtEmpPhoto.Text
    rsEmp.Fields("EmpPhoto").Value = StrStream.Read
    
    rsEmp.Update
    
    MsgBox "Employee Image Save Successfully", vbInformation
'    picEmp.Picture = Nothing
'    txtEmpPhoto.Text = ""
    cmdSaveEmp.Enabled = False
    Exit Sub
    
procNoPicture:
        MsgBox "Employee Image File Not Saved", vbExclamation
End Sub

Private Sub cmdSaveSign_Click()
On Error GoTo procNoPicture
    Set StrStream = New ADODB.Stream
    StrStream.Type = adTypeBinary
    StrStream.Open
    Dim rsSign As New ADODB.Recordset
    If cboCode.Text = "" Then
        MsgBox "Select The Employee", vbExclamation, "Blank Entry ......."
        cboCode.SetFocus
        Exit Sub
    ElseIf txtSignPhoto.Text = "" Then
        If MsgBox("Do You Wnat To Save Blank Employee Sigh Photo", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        cmdSignPhoto.SetFocus
        Exit Sub
        End If
    End If
    rsSign.Open "Select * From EmployeePhoto Where empcode = '" & cboCode.Text & "'", ConMain, adOpenDynamic, adLockOptimistic
    If rsSign.RecordCount < 1 Then
        rsSign.AddNew
        rsSign.Fields!Empcode = cboCode.Text
    End If
    
    If txtSignPhoto.Text <> "" Then StrStream.LoadFromFile txtSignPhoto.Text
    rsSign.Fields("SignPhoto").Value = StrStream.Read
       
    rsSign.Update
    MsgBox "Employee Signature Image Save Successfully", vbInformation
'    txtSignPhoto.Text = ""
'    PicSign.Picture = Nothing
    cmdSaveSign.Enabled = False
    
    Exit Sub
procNoPicture:
    Resume Next
    MsgBox "Employee Signature Image Save Successfully", vbInformation
End Sub

Private Sub cmdSignPhoto_Click()
    On Error GoTo Err
    ComDilog.Filter = "Pictures Files (*.Jpg; *.bmp)| *.Jpg;*.bmp"
    ComDilog.ShowOpen
    If ComDilog.FileName = "" Then Exit Sub
    txtSignPhoto.Text = ComDilog.FileName
    PicSign.Picture = LoadPicture(txtSignPhoto.Text)
    cmdSaveSign.Enabled = True
Err:
    If Err.Number = 481 Then
       MsgBox "Invalid Photo.....", vbExclamation, "Error Loading....."
       txtSignPhoto.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Call SetFormIcon(Me)
    Call ComboFill(cboCode, 1, 2)
    cmdSaveEmp.Enabled = False
    cmdSaveSign.Enabled = False
End Sub


Public Function LoadPictureFromDB(rsPic As ADODB.Recordset)
    On Error GoTo procNoPicture
    
    'If Recordset is Empty, Then Exit
    
    Set StrStream = New ADODB.Stream
    StrStream.Type = adTypeBinary
    StrStream.Open
    
    If Not IsNull(rsPic.Fields("EmpPhoto").Value) Then
        StrStream.Write rsPic.Fields("EmpPhoto").Value
        StrStream.SaveToFile "C:\Emp.bmp", adSaveCreateOverWrite
        picEmp.Picture = LoadPicture("C:\Emp.bmp")
        txtEmpPhoto.Text = "C:\Emp.bmp"
    End If
    
    If Not IsNull(rsPic.Fields("SignPhoto").Value) Then
        StrStream.Write rsPic.Fields("SignPhoto").Value
        StrStream.SaveToFile "C:\Sign.bmp", adSaveCreateOverWrite
        PicSign.Picture = LoadPicture("C:\Sign.bmp")
        txtSignPhoto.Text = "C:\Sign.bmp"
    End If
    
    LoadPictureFromDB = True
    Exit Function
    
procNoPicture:
    LoadPictureFromDB = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err
    
    Kill ("C:\Emp.bmp")
    Kill ("C:\Sign.bmp")
    
Err:
    If Err.Number = "53" Then Resume Next
End Sub



