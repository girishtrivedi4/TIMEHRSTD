VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImport 
   Caption         =   "Imports Data"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleMode       =   0  'User
   ScaleWidth      =   17166.71
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImports 
      Cancel          =   -1  'True
      Caption         =   "Import"
      Height          =   435
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      Height          =   435
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   5280
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame frSap 
      Height          =   795
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   6885
      Begin VB.TextBox txtFrom 
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
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "D"
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txtTo 
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
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "D"
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblFrPeri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import for the period from"
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
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   2460
      End
      Begin VB.Label lblToPeri 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   4590
         TabIndex        =   3
         Top             =   300
         Width           =   300
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7680
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "lblPath"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   585
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdImports_Click()
    CD.FileName = ""
    CD.Filter = "*.Dat|*.Dat"
    CD.Flags = cdlOFNFileMustExist
    CD.ShowSave
    If Trim(CD.FileName) = "" Then
        lblPath.Caption = ""
    End If
    If CD.FileName = "" Then
        MsgBox "Import File is not given", vbCritical
        Exit Sub
    End If
    lblPath.Caption = CD.FileName
End Sub

Private Sub cmdUpload_Click()
    If CD.FileName = "" Then
        MsgBox "Plese Select The Importing Dat File", vbExclamation, "Data Importing"
        Exit Sub
    End If
    If Not (IsDate(txtFrom.Text) Or IsDate(txtTo.Text)) Then
        MsgBox "Date Not Selected Properly", vbExclamation, "Data Importing"
        Exit Sub
    End If
 If ImportTextFile(ConMain, "DeviceLog", CD.FileName, vbTab) Then Exit Sub
End Sub

Private Sub Form_Load()
    txtFrom.Text = DateDisp(CStr(Date))
    txtTo.Text = DateDisp(CStr(Date))
    lblPath.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandler
    If FindTable("DeviceLogTemp") Then ConMain.Execute "Drop Table DeviceTemp"
    ConMain.Execute "SELECT Distinct * INTO DeviceTemp FROM DeviceLog"
    If FindTable("DeviceLog") Then ConMain.Execute "Drop Table DeviceLog"
    ConMain.Execute "SELECT * INTO DeviceLog FROM DeviceTemp"
    If FindTable("DeviceLog") Then ConMain.Execute "Drop Table DeviceTemp"
    Exit Sub
errHandler:
    WriteLog (Err.Description)
End Sub

Private Sub txtFrom_Click()
    Call ShowCalendar
End Sub

Private Sub txtTo_Click()
    Call ShowCalendar
End Sub

Public Function ImportTextFile(cn As Object, _
  ByVal tblName As String, FileFullPath As String, _
  Optional FieldDelimiter As String = ",", _
  Optional RecordDelimiter As String = vbCrLf) As Boolean

Dim sFileContents As String
Dim iFileNum As Integer
Dim sTableSplit() As String
Dim lCtr As Integer
Dim iCtr As Integer
Dim lRecordCount As Long
Dim Punches() As String
Dim sSQL As String
Dim Counter As Integer
Dim Consql As New ADODB.Connection



On Error GoTo errHandler

Consql.ConnectionString = cn.ConnectionString
Consql.Open

'Consql.BeginTrans

If Not TypeOf cn Is ADODB.Connection Then Exit Function
If Dir(FileFullPath) = "" Then Exit Function

If Consql.State = 0 Then Consql.Open

iFileNum = FreeFile

Open FileFullPath For Input As #iFileNum
sFileContents = Input(LOF(iFileNum), #iFileNum)
Close #iFileNum
'split file contents into rows
sTableSplit = Split(sFileContents, vbCrLf)
    
lRecordCount = UBound(sTableSplit)

Consql.BeginTrans

For lCtr = 0 To lRecordCount - 1
     
    Punches = Split(sTableSplit(lCtr), vbTab)
        If Format(Trim(Punches(1)), "dd/mm/yyyy") >= DateCompDate(txtFrom.Text) And Format(Trim(Punches(1)), "dd/mm/yyyy") <= DateCompDate(txtTo.Text) Then
            sSQL = "INSERT INTO DeviceLog Values( '" & Trim(Punches(0)) & "','" & CDate(Trim(Punches(1))) & "')"
            cn.Execute sSQL
            Me.Caption = "Data Importing " & sTableSplit(lCtr)
            Me.Refresh
            Counter = Counter + 1
        End If
Next lCtr
Me.Caption = "Imports Data"
Consql.CommitTrans
Set Consql = Nothing

Close #iFileNum
    MsgBox "Transaction " & Counter & " Importing Successfully", vbInformation, "Data I1mporting"
ImportTextFile = True

Exit Function
'
errHandler:

MsgBox Err.Description

If Consql.State <> 0 Then Consql.RollbackTrans
If iFileNum > 0 Then Close #iFileNum
Set Consql = Nothing

End Function


