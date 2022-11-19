VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmF12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Electronics"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6780
      TabIndex        =   4
      Top             =   5010
      Width           =   1545
   End
   Begin VB.CommandButton cmdZap 
      Caption         =   "&Zap"
      Height          =   375
      Left            =   5250
      TabIndex        =   3
      Top             =   5010
      Width           =   1545
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   4605
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   8123
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADC1 
      Height          =   375
      Left            =   90
      Top             =   990
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cboTab 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "frmF12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboTab_Click()
On Error GoTo ERR_P
If cboTab.Text = "" Then Exit Sub
ADC1.RecordSource = Trim(cboTab.Text)
ADC1.Refresh
'Sleep (5000)
DG1.Refresh
cmdZap.Enabled = True
Exit Sub
ERR_P:
    ShowError ("cboTable :: " & Me.Caption)
    cmdZap.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdZap_Click()
On Error GoTo ERR_P
If Trim(cboTab.Text) = "" Then Exit Sub
If MsgBox("This Will Delete All the Records from the Table '" & cboTab.Text & "'" _
, vbQuestion + vbYesNo) = vbYes Then
    Call TruncateTable(cboTab.Text)
    Sleep (5000)
    ADC1.Refresh
    DG1.ClearFields
    DG1.Refresh
End If
Exit Sub
ERR_P:
    ShowError ("cmdZap :: " & Me.Caption)
End Sub

Private Sub DG1_Click()
''
End Sub

Private Sub DG1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 4      '' DEncrypt
        bytFormToLoad = 7
        PassFrm.Show vbModal
    Case 16     '' PUser
        bytFormToLoad = 8
        PassFrm.Show vbModal
End Select
End Sub

Private Sub AddTables()         '' Add User Tables to the List
On Error GoTo ERR_P
Dim adTmp As New ADODB.Recordset, intTmp As Integer
adTmp.CursorType = adOpenStatic
adTmp.LockType = adLockOptimistic
Select Case bytBackEnd
    Case 1, 2
        Set adTmp = VstarDataEnv.cnDJConn.OpenSchema(adSchemaTables)
        cboTab.Clear
        If adTmp.RecordCount = 0 Then Exit Sub
        For intTmp = 0 To adTmp.RecordCount - 1
            If adTmp(3) = "TABLE" Then cboTab.AddItem UCase(adTmp(2))
            adTmp.MoveNext
        Next
    Case 3
        Set adTmp = VstarDataEnv.cnDJConn.Execute("select UPPER(Table_Name) from Tabs ")
        If adTmp.RecordCount = 0 Then Exit Sub
        For intTmp = 0 To adTmp.RecordCount - 1
            cboTab.AddItem UCase(adTmp(0))
            adTmp.MoveNext
        Next
End Select
Exit Sub
ERR_P:
    ShowError (" Add Tables :: " & Me.Caption)
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me)
If Not ADCConnect Then
    MsgBox "Database Connection Failed", vbCritical, App.EXEName
    DG1.Visible = False
    cmdZap.Visible = False
End If
ADC1.Visible = False
Call AddTables
End Sub

Private Function ADCConnect() As Boolean
On Error GoTo ERR_P
Set DG1.DataSource = ADC1
ADC1.ConnectionString = VstarDataEnv.cnDJConn.ConnectionString & strDBPass
ADC1.CommandType = adCmdTable
ADCConnect = True
Exit Function
ERR_P:
    ShowError ("ADCConnect :: " & Me.Caption)
End Function

Private Sub Form_Unload(Cancel As Integer)
ADC1.ConnectionString = ""
End Sub
