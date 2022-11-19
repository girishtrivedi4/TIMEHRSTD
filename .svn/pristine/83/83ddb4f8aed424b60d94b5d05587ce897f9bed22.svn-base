VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDSN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Connection"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdWiz 
      Caption         =   "&Show System DSN Wizard"
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
      Left            =   840
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCre 
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2880
      TabIndex        =   3
      Top             =   3360
      Width           =   1185
   End
   Begin TabDlg.SSTab STabDSN 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3889
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Select Link"
      TabPicture(0)   =   "frmDSN.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSlDSN"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Create Link"
      TabPicture(1)   =   "frmDSN.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "txtCrDataBaseName"
      Tab(1).Control(2)=   "cmdCrLink"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Delete Link"
      TabPicture(2)   =   "frmDSN.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdDlLink"
      Tab(2).Control(1)=   "txtDlLink"
      Tab(2).Control(2)=   "Label5"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdCrLink 
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -73320
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdDlLink 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -73440
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox txtDlLink 
         Height          =   315
         Left            =   -72930
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   3570
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCrDataBaseName 
         Height          =   285
         Left            =   -72930
         TabIndex        =   4
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox txtSlDSN 
         Height          =   315
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Data Base Link"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74820
         TabIndex        =   29
         Top             =   765
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base Link Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74775
         TabIndex        =   23
         Top             =   765
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Data Base Link"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   765
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   28
      Top             =   1920
      Width           =   6045
      Begin VB.TextBox txtDSN 
         Height          =   285
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5220
         TabIndex        =   21
         Top             =   1110
         Width           =   345
      End
      Begin VB.TextBox txtSer 
         Height          =   285
         Left            =   3810
         TabIndex        =   17
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox txtPass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   870
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1140
         Width           =   1755
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Left            =   870
         TabIndex        =   11
         Top             =   720
         Width           =   1755
      End
      Begin VB.ComboBox cboBack 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1725
      End
      Begin VB.ComboBox cboVersion 
         Height          =   315
         ItemData        =   "frmDSN.frx":0054
         Left            =   3780
         List            =   "frmDSN.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1125
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DSN Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2800
         TabIndex        =   14
         Top             =   300
         Width           =   810
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P&ath"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2800
         TabIndex        =   18
         Top             =   1170
         Width           =   330
      End
      Begin VB.Label lblSer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server &Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2805
         TabIndex        =   16
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   45
         TabIndex        =   10
         Top             =   765
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Back End"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "IV SOFTTECH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1305
      Left            =   0
      TabIndex        =   27
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   6045
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   5760
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Con As New ADODB.Connection

Private Function CreateDSN(BackEnd As Byte) As Boolean
On Error GoTo ERR_P
Dim DataSourceName As String
Dim DatabaseName As String
Dim Description As String
Dim DriverPath As String
Dim DriverName As String
Dim LastUser As String
Dim Server As String
Dim Password As String
Dim lResult As Long
Dim hKeyHandle As Long
DriverPath = Space(255)
lResult = GetSystemDirectory(DriverPath, Len(DriverPath))
DriverPath = Trim(DriverPath)
DriverPath = DriverPath & "\ODBCJT32.DLL"
'Specify the DSN parameters.
Select Case BackEnd
    Case 1  '' MS-SQL Server
        DataSourceName = TDSN.DSNName
        DatabaseName = TDSN.Database
        Description = TDSN.DSNName
        LastUser = DEncryptDat(TDSN.UserName, 1) & Chr(0)
        Server = TDSN.ServerName
        DriverName = "SQL Server"
        Password = DEncryptDat(TDSN.Password, 1) & Chr(0)
        
        lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & TDSN.DSNName, hKeyHandle)
        lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, ByVal Server, Len(Server))
        lResult = RegSetValueEx(hKeyHandle, "DATABASE", 0&, REG_SZ, ByVal DatabaseName, Len(DatabaseName))
        lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, ByVal Description, Len(Description))
        lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, ByVal DriverName, Len(DriverPath))
        lResult = RegSetValueEx(hKeyHandle, "UID", 0&, REG_SZ, ByVal LastUser, Len(LastUser))
        lResult = RegSetValueEx(hKeyHandle, "PWD", 0&, REG_SZ, ByVal Password, Len(Password))
        lResult = RegCloseKey(hKeyHandle)
    Case 2  '' MS-Access
        DataSourceName = TDSN.DSNName
        DatabaseName = TDSN.path
        Description = TDSN.DSNName
        LastUser = ""
        Server = TDSN.ServerName
        DriverName = "Microsoft Access Driver (*.MDB)"
        Password = TDSN.Password
        
        lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & TDSN.DSNName, hKeyHandle)
        lResult = RegSetValueEx(hKeyHandle, "DBQ", 0&, REG_SZ, ByVal DatabaseName, Len(DatabaseName))
        lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, ByVal Description, Len(Description))
        lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, ByVal DriverPath, Len(DriverPath))
        lResult = RegSetValueEx(hKeyHandle, "FIL", 0&, REG_SZ, ByVal Server, Len(Server))
        lResult = RegSetValueEx(hKeyHandle, "PWD", 0&, REG_SZ, ByVal Password, Len(Password))
        lResult = RegSetValueEx(hKeyHandle, "DriverId", 0&, REG_DWORD, 25&, 4)
        lResult = RegSetValueEx(hKeyHandle, "Safe Transactions", 0&, REG_DWORD, 0&, 4)
        lResult = RegSetValueEx(hKeyHandle, "UID", 0&, REG_SZ, ByVal LastUser, 0)
        lResult = RegCloseKey(hKeyHandle)
        
   Case 3  '' Oracle
        DataSourceName = TDSN.DSNName
        DatabaseName = TDSN.Database
        Description = TDSN.DSNName
        LastUser = DEncryptDat(TDSN.UserName, 1) & Chr(0)
        Server = TDSN.ServerName
        '
        If cboVersion.ListIndex = 0 Then
            DriverName = "Microsoft ODBC for Oracle"
        ElseIf cboVersion.ListIndex = 1 Then
            DriverName = "Oracle in Oracle9i"
        ElseIf cboVersion.ListIndex = 2 Then
            DriverName = "Oracle in Orahome92"
        End If
        DriverPath = "E:\Oracle\Ora81\BIN\SQORA32.DLL"
        Password = DEncryptDat(TDSN.Password, 1) & Chr(0)
        
        lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & TDSN.DSNName, hKeyHandle)
        lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, ByVal Server, Len(Server))
        lResult = RegSetValueEx(hKeyHandle, "DATABASE", 0&, REG_SZ, ByVal DatabaseName, Len(DatabaseName))
        lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, ByVal Description, Len(Description))
        lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, ByVal DriverName, Len(DriverPath))
        lResult = RegSetValueEx(hKeyHandle, "UID", 0&, REG_SZ, ByVal LastUser, Len(LastUser))
        lResult = RegSetValueEx(hKeyHandle, "PWD", 0&, REG_SZ, ByVal Password, Len(Password))
        lResult = RegCloseKey(hKeyHandle)
End Select
'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
'Specify the new value.
lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, ByVal DriverName, Len(DriverName))
'Close the key.
lResult = RegCloseKey(hKeyHandle)
''commented by SG07 as per kk's requirement
''Call WriteToINI
''
CreateDSN = True
Exit Function
ERR_P:
    ShowError ("CreateDSN::")
End Function

Private Function DEncryptDat(pwd As String, bytTimes As Byte) As String
Dim pwdStr As String, i As Integer
For i = 1 To Len(pwd)
        pwdStr = pwdStr & Chr(Asc(Mid(pwd, i, 1)) Xor IIf(bytTimes = 1, 11, 12))
Next i
DEncryptDat = pwdStr
End Function

Private Function ExistingDSN() As Boolean
ExistingDSN = True
End Function

Private Function ValidateDSN() As Boolean
On Error GoTo ERR_P
Dim blnTmp As Boolean
Select Case TDSN.BackEnd
    Case 1      '' MS-SQL Server
        ''If Trim(txtSer.Text) = "" Then
        ''    MsgBox "Please Enter ServerName", vbCritical: txtSer.SetFocus: Exit Function
        ''End If
        If Trim(txtUser.Text) = "" Then
            MsgBox "Please Enter UserName", vbCritical: txtUser.SetFocus: Exit Function
        End If
        If Trim(txtPass.Text) = "" Then
            MsgBox "Please Enter Password", vbCritical: txtPass.SetFocus: Exit Function
        End If
        TDSN.UserName = Trim(txtUser.Text)
        TDSN.Password = Trim(txtPass.Text)
        TDSN.ServerName = Trim(txtSer.Text)
        TDSN.Database = Dbname
        TDSN.DSNName = Trim(txtDSN.Text)
        TDSN.path = ""
    Case 2      '' Ms-Access
        If Trim(txtPass.Text) = "" Then
            MsgBox "Please Enter Password", vbCritical: txtPass.SetFocus: Exit Function
        End If
        If Trim(txtPath.Text) = "" Then
            MsgBox "Please Enter Database Path", vbCritical: txtPath.SetFocus: Exit Function
        End If
        TDSN.UserName = Trim(txtUser.Text)
        TDSN.Password = Trim(txtPass.Text)
        TDSN.ServerName = "MS Access"
        TDSN.Database = Trim(txtPath.Text)
        TDSN.DSNName = Trim(txtDSN.Text)
        TDSN.path = Trim(txtPath.Text)
    Case 3      '' Oracle
        ''If Trim(txtSer.Text) = "" Then
        ''    MsgBox "Please Enter ServerName", vbCritical: txtSer.SetFocus: Exit Function
        ''End If
        If Trim(txtUser.Text) = "" Then
            MsgBox "Please Enter UserName", vbCritical: txtUser.SetFocus: Exit Function
        End If
        If Trim(txtPass.Text) = "" Then
            MsgBox "Please Enter Password", vbCritical: txtPass.SetFocus: Exit Function
        End If
        TDSN.UserName = Trim(txtUser.Text)
        TDSN.Password = Trim(txtPass.Text)
        TDSN.ServerName = Trim(txtSer.Text)
        TDSN.Database = Trim(txtUser.Text)
        TDSN.DSNName = Trim(txtDSN.Text)
        TDSN.path = ""
End Select
blnTmp = ExistingDSN
ValidateDSN = blnTmp
Exit Function
ERR_P:
    ShowError ("ValidateDSN::")
End Function

Private Sub cboBack_Click()
TDSN.BackEnd = cboBack.ListIndex + 1
Select Case TDSN.BackEnd
    Case 1      '' MS-SQL Server
        lblSer.Visible = True
        txtSer.Visible = True
        lblPath.Visible = False
        txtPath.Visible = False
        cmdPath.Visible = False
        lblUser.Visible = True
        txtUser.Text = InVar.strUser
        txtUser.Visible = True
        txtPass.Text = InVar.strPass
        '
        cboVersion.Visible = False
       Case 2      '' MS-Access
        lblSer.Visible = False
        txtSer.Visible = False
        lblPath.Visible = True
        txtPath.Visible = True
        cmdPath.Visible = True
        lblUser.Visible = False
        txtUser.Text = ""
        txtUser.Visible = False
'        txtPass.Text = InVar.strPass
        '
        cboVersion.Visible = False
        lblPath.Caption = " Path"
    Case 3      '' Oracle
        lblSer.Visible = True
        txtSer.Visible = True
        lblPath.Visible = False
        txtPath.Visible = False
        cmdPath.Visible = False
        lblUser.Visible = True
        txtUser.Text = InVar.strUser
        txtUser.Visible = True
        txtPass.Text = InVar.strPass
        '
        cboVersion.Visible = True
        lblPath.Caption = " Version"
        '
        Label3.Top = 1200
        txtPass.Top = 1140
        lblPath.Top = 1200
        cmdPath.Top = 1140
        txtPath.Top = 1140
End Select
        lblNotes.Caption = "Due to Some Reasons the database connection may be Corrupted or Deleted. " & _
        "Connection can be Created with a Valid User Name,Password (Case Sensitive) and a Server Name. " & _
        vbCrLf & "Please Contact Your System Administrator for further Details."

End Sub

Private Sub cmdCre_Click()
On Error GoTo ERR_P:
If GetFlagStatus("OPTIONALDSN") Then
    If STabDSN.Tab = 1 Then
        If adrsTemp.RecordCount > 0 Then
        adrsTemp.MoveFirst
        adrsTemp.Find " ConName = '" & Trim(txtSlDSN.Text) & "'"
        End If
        If Not adrsTemp.EOF Then
            MsgBox "Data Base Link Name Already Exists, Please Enter Another Link Name", vbExclamation, "Duplicate Entry"
            txtCrDataBaseName.SetFocus
            Exit Sub
        End If
        
        Dim TempString As String
        TempString = GetConnString
                
        ConMain.ConnectionString = TempString
        ConMain.Open
        
        
        Con.Execute "Insert Into ConString VALUES('" & Trim(txtCrDataBaseName.Text) & "','" & Trim(TempString) & "')"
        Unload Me
        Exit Sub
    End If
Else
    If Not ValidateDSN Then Exit Sub
End If
    Call MakeDSN
Exit Sub
ERR_P:
    'Resume Next
    If Err.Number = -2147467259 Or -2147217843 Then
        MsgBox "Invalied Server Name Or Data Base Name Or Server Connection Not Found ......", vbExclamation, "Error Connecting"
        End
    End If
    ShowError (Err.Description & " :: " & Me.Caption)
End Sub

    Private Sub cmdCrLink_Click()
    On Error GoTo Err
    Dim TempString As String
    If Trim(txtCrDataBaseName.Text) = "" Then
        MsgBox "Enter The Proper Connection Name", vbExclamation, "Blank Entry..."
        txtCrDataBaseName.SetFocus
        Exit Sub
    End If
    If adrsTemp.RecordCount > 0 Then
    adrsTemp.MoveFirst
    adrsTemp.Find " ConName = '" & Trim(txtCrDataBaseName.Text) & "'"
        If Not adrsTemp.EOF Then
            MsgBox "Data Base Link Name Already Exists, Please Enter Another Link Name", vbExclamation, "Duplicate Entry"
            txtCrDataBaseName.SetFocus
            Exit Sub
        End If
    End If
    TempString = GetConnString
    ConMain.ConnectionString = TempString
    ConMain.Open
    If InVar.strSer = 1 Then
        If UCase(ConMain.DefaultDatabase) = "MASTER" Or ConMain.DefaultDatabase = "" Then
            MsgBox "Data Base Not Selected Properly, Please Recreate the Link", vbInformation, "Error link...."
            End
        End If
    End If

    ConLog.Execute "Insert Into ConString VALUES('" & Trim(txtCrDataBaseName.Text) & "','" & Trim(TempString) & "')"
    MsgBox "New Data Base Link Created Succefully", vbInformation
    Unload Me
    Exit Sub
Err:
    MsgBox "Error Connecting to the Database :: Cannot Proceed" & _
    vbCrLf & "Contact Your System Administrator", vbCritical, strVersionWithTital
    End
End Sub

Private Sub cmdDlLink_Click()
    If txtDlLink.Text = "" Then
        MsgBox "Select The Proper Connection From Deletion", vbExclamation, "Blank Entry..."
        txtDlLink.SetFocus
        Exit Sub
    End If
    If MsgBox("Do You Sure Want To Delete Database Link", vbYesNo + vbQuestion) = vbYes Then
       ConLog.Execute "Delete From ConString Where Conname = '" & txtDlLink.Text & "'"
        MsgBox "Link Deleted Successfully", vbInformation
    End If
    Call FillConCombos
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err
'    TDSN.DSNName = txtSlDSN.Text
    If txtSlDSN.Text = "" Then
        MsgBox "Select The Proper Connection", vbExclamation, "Blank Entry..."
        txtSlDSN.SetFocus
        Exit Sub
    End If
    If adrsTemp.RecordCount > 0 Then
    adrsTemp.MoveFirst
    adrsTemp.Find " ConName = '" & txtSlDSN.Text & "'"
    End If
    ConMain.ConnectionString = adrsTemp.Fields(1)
    ConMain.Open
    ConMain.CursorLocation = adUseClient
    If InVar.strSer = 1 Then
        If UCase(ConMain.DefaultDatabase) = "MASTER" Or ConMain.DefaultDatabase = "" Then
            MsgBox "Data Base Not Selected Properly, Please Recreate the Link", vbInformation, "Error link...."
            Exit Sub
        End If
    End If
    
    Unload Me
    Exit Sub
Err:
        MsgBox "Error Connecting to the Database :: Cannot Proceed" & _
        vbCrLf & "Contact Your System Administrator", vbCritical, strVersionWithTital
        End
End Sub

Private Sub txtSlDSN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub txtCrDataBaseName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdCre_Click
End Sub

Private Sub STabDSN_Click(PreviousTab As Integer)
    If STabDSN.Tab = 0 Then
        cmdCre.Enabled = False
    Else
        If STabDSN.Tab = 1 Then
            cmdCre.Enabled = True
'            txtCrUserName.SetFocus
        End If
        
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPath_Click()
On Error Resume Next
With CDLG
    If txtPath.Text <> "" Then .InitDir = Replace(txtPath.Text, "\" & CDLG.FileTitle, "")
    .ShowOpen
End With
If CDLG.FileName <> "" Then txtPath.Text = CDLG.FileName
End Sub

Private Sub cmdWiz_Click()
On Error Resume Next
Call Shell("ODBCAD32.EXE", vbNormalFocus)
If Err.Number <> 0 Then MsgBox Err.Description & vbCrLf & "Unable to Show the System DSN Wizard"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 28 Then
    bytFormToLoad = 10
    PassFrm.Show vbModal
    bytFormToLoad = 0
    End
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SetFormIcon(Me)
'' Fill Back End Combo
With cboBack
    .AddItem "1  MS-SQL Server"
    .AddItem "2  MS-Access"
    .AddItem "3  Oracle"
    .ListIndex = IIf(CByte(InVar.strSer) > 0 And CByte(InVar.strSer) <= 3, CByte(InVar.strSer) - 1, 1)
End With
txtDSN.Text = "VisualStarDSN"
With CDLG
    .DefaultExt = "*.MDB"
    .Filter = "*.MDB | *.MDB"
    .Flags = cdlOFNFileMustExist
End With

Me.Caption = "Database Connection"

    STabDSN.Tab = 0
'    Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data\VSTARREC.MDB;Persist Security Info=False;Jet OLEDB:Database Password=ATTENDOLOG"
'    Con.Open
'    Con.CursorLocation = adUseClient
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "Select * from MSysObjects where Name = 'ConString'", ConLog, adOpenKeyset, adLockReadOnly
    
    If adrsTemp.RecordCount <= 0 Then
        Con.Execute "Create table ConString(ConName text (20), ConString Text(200))"
    End If
   Call FillConCombos

End Sub

Private Sub MakeDSN()
If CreateDSN(TDSN.BackEnd) Then
    MsgBox "DSN Created Successfully"
Else
    MsgBox "DSN Creation Failed"
End If
Unload Me
End Sub

Private Sub txtDSN_DblClick()
txtDSN.Locked = Not txtDSN.Locked
End Sub

 Public Function GetConnString() As String
    Dim objDataLink As New DataLinks
    Dim strConn As String
    On Error GoTo GetConnString_Error
    
    objDataLink.hwnd = Me.hwnd
    strConn = objDataLink.PromptNew
    GetConnString = strConn
    
    
    On Error GoTo 0
    Exit Function
    
GetConnString_Error:
    If Err.Number = 91 Then Exit Function
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetConnString of Form frmDataExportSetup"

End Function

Private Sub FillConCombos() ' 01-10-09
    txtSlDSN.clear
    txtDlLink.clear
   If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "Select * From ConString", ConLog, adOpenKeyset, adLockReadOnly
       If adrsTemp.RecordCount > 0 Then
        For i = 1 To adrsTemp.RecordCount
            txtSlDSN.AddItem adrsTemp.Fields(0)
            txtDlLink.AddItem adrsTemp.Fields(0)
            adrsTemp.MoveNext
        Next
    End If
End Sub


Private Sub ConnectDB()
    If adrsTemp.State = 1 Then adrsTemp.Close
    adrsTemp.Open "Select * From ConString", Con, adOpenDynamic, adLockOptimistic
    If adrsTemp.RecordCount > 0 Then
        ConMain.ConnectionString = adrsTemp.Fields(1)
        ConMain.Open
    End If
    If InVar.strSer = 1 Then
        If UCase(ConMain.DefaultDatabase) = "MASTER" Or ConMain.DefaultDatabase = "" Then
            MsgBox "Data Base Not Selected Properly, Please Recreate the Link", vbInformation, "Error link...."
            Exit Sub
        End If
    End If
End Sub


