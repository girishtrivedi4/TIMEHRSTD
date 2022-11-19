VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Version 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Okcmd 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   4110
      Width           =   1215
   End
   Begin VB.Label VersionLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Left            =   2880
      TabIndex        =   0
      Top             =   60
      Width           =   570
   End
   Begin MSForms.ListBox FileDetails 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   6375
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "11245;6298"
      ColumnCount     =   2
      MatchEntry      =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Version"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub Form_Activate()
Dim VersLength As Integer
VersLength = VersionLbl.Width
VersionLbl.Move (Version.Width / 2 - VersLength / 2)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)
Call RetCaptions
FileDetails.ColumnCount = 2
FileDetails.ColumnWidths = "3cm, 7cm"
Dim Temparray(11, 1) As String
Temparray(0, 0) = NewCaptionTxt("57003", adrsC)                 ''Comments
Temparray(0, 1) = App.Comments
Temparray(1, 0) = NewCaptionTxt("57004", adrsC)                 '' Company Name
Temparray(1, 1) = "IV SOFTTECH" 'App.CompanyName
Temparray(2, 0) = NewCaptionTxt("57005", adrsC)                 '' File description
Temparray(2, 1) = App.FileDescription
Temparray(3, 0) = NewCaptionTxt("57006", adrsC)                 '' File version
Temparray(3, 1) = App.Major & "." & App.Minor & "." & App.Revision
Temparray(4, 0) = NewCaptionTxt("57007", adrsC)                 '' Internal Name
Temparray(4, 1) = App.Title
'Temparray(5, 0) = NewCaptionTxt("57008", adrsC)                 '' Legal Copyright
'Temparray(5, 1) = App.LegalCopyright
'Temparray(6, 0) = NewCaptionTxt("57009", adrsC)                 '' Legal trademarks
'Temparray(6, 1) = App.LegalTrademarks
Temparray(5, 0) = NewCaptionTxt("57010", adrsC)                 '' original filename
Temparray(5, 1) = App.EXEName
Temparray(6, 0) = NewCaptionTxt("57011", adrsC)                 '' product name
Temparray(6, 1) = App.ProductName
Temparray(7, 0) = NewCaptionTxt("57012", adrsC)                 '' Product version
Temparray(7, 1) = App.Major & "." & App.Minor & "." & App.Revision
Temparray(8, 0) = NewCaptionTxt("57013", adrsC)                '' special build for
        FileDetails.List = Temparray

Erase Temparray
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub Okcmd_Click()
    Unload Me
End Sub

Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '57%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = "TimeHR Version"
VersionLbl.Caption = NewCaptionTxt("57002", adrsC)
OKCmd.Caption = NewCaptionTxt("00002", adrsMod)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub
