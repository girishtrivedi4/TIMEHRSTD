VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmF10 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   5910
      Width           =   1725
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   5865
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   10345
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
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
         DataField       =   "CaptEng"
         Caption         =   "English"
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
         DataField       =   "CaptOther"
         Caption         =   "Other Language"
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
            ColumnAllowSizing=   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   4694.74
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   4529.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADC1 
      Height          =   375
      Left            =   30
      Top             =   5490
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
End
Attribute VB_Name = "frmF10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub DG1_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Call SetFormIcon(Me, True)
Me.Caption = "Edit Language Captions"
If Not ADCConnect Then
    MsgBox "Database Connection Failed", vbCritical, App.EXEName
    DG1.Visible = False
End If
ADC1.Visible = False
End Sub

Private Function ADCConnect() As Boolean
On Error GoTo ERR_P
Set DG1.DataSource = ADC1
ADC1.ConnectionString = VstarDataEnv.cnDJConn.ConnectionString & strDBPass
ADC1.CommandType = adCmdText
If strFormCommand = "0" Then
    ADC1.RecordSource = "Select CaptEng,CaptOther from NewCaptions Where Id " & _
        "Like '00%' Or ID Like 'M%' order by ID"
Else
    ADC1.RecordSource = "Select CaptEng,CaptOther from NewCaptions Where Id " & _
        "Like '" & strFormCommand & "%' order by ID"
End If
ADC1.Refresh
DG1.Refresh
ADCConnect = True
Exit Function
ERR_P:
    ShowError ("ADCConnect :: " & Me.Caption)
End Function

Private Sub Form_Unload(Cancel As Integer)
ADC1.ConnectionString = ""
End Sub
