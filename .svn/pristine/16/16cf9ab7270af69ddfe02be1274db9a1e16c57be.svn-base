VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Encrypt Decrypt"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LockControls    =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRel 
      Height          =   405
      Left            =   9060
      TabIndex        =   9
      Top             =   0
      Width           =   1965
   End
   Begin RichTextLib.RichTextBox txtDec 
      Height          =   3885
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   6853
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmmain.frx":0000
   End
   Begin VB.CommandButton cmdSaveDec 
      Height          =   375
      Left            =   5820
      TabIndex        =   2
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdSaveEnc 
      Height          =   375
      Left            =   5820
      TabIndex        =   6
      Top             =   4350
      Width           =   1155
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9540
      TabIndex        =   8
      Top             =   4320
      Width           =   1485
   End
   Begin VB.TextBox txtFEnc 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4350
      Width           =   4635
   End
   Begin VB.CommandButton cmdEnc 
      Height          =   375
      Left            =   30
      TabIndex        =   4
      Top             =   4350
      Width           =   1095
   End
   Begin VB.TextBox txtFDec 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   4635
   End
   Begin VB.CommandButton cmdDec 
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7260
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtEnc 
      Height          =   3885
      Left            =   0
      TabIndex        =   7
      Top             =   4740
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   6853
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmmain.frx":0082
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDec_Click()
With CD1
    .FileName = ""
    .Filter = "*.Script Files|*.SCP"
    .Flags = cdlOFNFileMustExist
    .ShowOpen
    txtFDec.Text = .FileName
End With
If Trim(txtFDec.Text) = "" Then Exit Sub
txtDec.Text = DecryptSCR(OpenFile(txtFDec.Text))
txtDec.SaveFile App.Path & "\Tmp.TXT"
txtDec.LoadFile App.Path & "\Tmp.TXT"
txtDec.Text = Right(txtDec.Text, Len(txtDec.Text) - 4)
Kill App.Path & "\TMP.TXT"
End Sub

Private Sub cmdEnc_Click()
With CD1
    .FileName = ""
    .Filter = "*.Text Files|*.TXT"
    .Flags = cdlOFNFileMustExist
    .ShowOpen
    txtFEnc.Text = .FileName
End With
If Trim(txtFEnc.Text) = "" Then Exit Sub
txtEnc.Text = EncryptSCR(OpenFile(txtFEnc.Text))
''Call cmdSaveEnc_Click
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Function OpenFile(strFileName As String) As String
On Error GoTo ERR_P
Dim strDataL As String, strTmp As String
If Dir(strFileName) = "" Then
    MsgBox "File " & strFileName & " not found", vbCritical, App.EXEName
    Exit Function
End If
Open strFileName For Input As #2          '' Open the File
strTmp = ""
Do While Not EOF(2)
    Line Input #2, strDataL             '' Get Record from the DAT File
    strTmp = strTmp & vbCrLf & strDataL
Loop
Close #2
OpenFile = strTmp
Exit Function
ERR_P:
    MsgBox "Error Opening File " & strFileName & vbCrLf & Err.Description, vbCritical
End Function


Private Sub cmdSaveDec_Click()
With CD1
    .FileName = ""
    .Filter = "*.Text Files|*.TXT"
    .Flags = cdlOFNOverwritePrompt
    .ShowSave
    If Trim(.FileName) = "" Then Exit Sub
    Call SaveFile(.FileName, txtDec.Text)
End With
End Sub

Private Sub SaveFile(strFileName As String, strText As String)
On Error GoTo ERR_P
Open strFileName For Output As #2               '' Open the File
Print #2, strText                               '' Save the File
Close #2
Exit Sub
ERR_P:
    MsgBox "Error Saving File " & strFileName & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub cmdSaveEnc_Click()
With CD1
    If Trim(.FileName) = "" Then Exit Sub
    .FileName = Left(.FileName, Len(.FileName) - 4)
    .Filter = "*.Script Files|*.SCP"
    .Flags = cdlOFNOverwritePrompt
    .ShowSave
    Call SaveFile(.FileName, txtEnc.Text)
End With
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me, True)
End Sub
