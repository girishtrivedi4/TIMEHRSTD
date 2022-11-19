VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSelectShift 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUndo 
      Caption         =   " "
      Height          =   405
      Left            =   3630
      MaskColor       =   &H80000004&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtShifts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
   Begin VB.CommandButton cmdReset 
      Height          =   375
      Left            =   3630
      TabIndex        =   3
      Top             =   1050
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Height          =   375
      Left            =   3630
      TabIndex        =   4
      Top             =   1650
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3630
      TabIndex        =   5
      Top             =   2250
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   450
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSelectShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    strRotPass = Trim(txtShifts.Text)
    
    strAutoG = Trim(txtShifts.Text)
    ''
    Unload Me
End Sub

Private Sub cmdReset_Click()
    txtShifts.Text = ""
End Sub

Private Sub cmdUndo_Click()
On Error Resume Next
Do
    If txtShifts.Text = "" Then Exit Do
    txtShifts.Text = Left(txtShifts.Text, Len(txtShifts.Text) - 1)
Loop While Right(txtShifts.Text, 1) <> "."
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdOK_Click
If KeyAscii = 8 Then Call cmdUndo_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SetFormIcon(Me, True)
Call RetCaptions
Call FillGrid
cmdUndo.Picture = LoadPicture(App.path & "\Images\Undo.Bmp")
End Sub

Private Sub RetCaptions()
On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '47%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("47001", adrsC)
With MSF1
    .TextMatrix(0, 0) = NewCaptionTxt("00031", adrsMod)
    .TextMatrix(0, 1) = NewCaptionTxt("47002", adrsC)
    .TextMatrix(0, 2) = NewCaptionTxt("47003", adrsC)
    .TextMatrix(0, 3) = NewCaptionTxt("47004", adrsC)
    .ColWidth(0) = MSF1.Width * (22 / 100)
    .ColWidth(1) = MSF1.Width * (22 / 100)
    .ColWidth(2) = MSF1.Width * (22 / 100)
    .ColWidth(3) = MSF1.Width * (22 / 100)
End With
cmdReset.Caption = "Reset" 'NewCaptionTxt("47005", adrsC)
cmdCan.Caption = "Cancel" ' NewCaptionTxt("00003", adrsMod)
cmdOk.Caption = "OK" ' NewCaptionTxt("00002", adrsMod)
txtShifts.Text = strRotPass
End Sub

Private Sub FillGrid()
On Error GoTo ERR_P
Dim bytTmp As Byte
If adRsintshft.State = 1 Then adRsintshft.Close
adRsintshft.Open "Select Shift,Shf_In,Shf_Out,Night from InstShft WHERE Shift <> '100'", _
ConMain, adOpenKeyset
MSF1.Rows = adRsintshft.RecordCount + 4

Dim bytT1 As Byte
If Right(strAutoG, 2) = "Em" Then
    MSF1.TextMatrix(1, 0) = "ALL"
    bytTmp = 2
    txtShifts.Text = Left(strAutoG, Len(strAutoG) - 2)
Else
    bytTmp = 1
End If
''
Do While Not adRsintshft.EOF
    With MSF1
        .TextMatrix(bytTmp, 0) = Trim(adRsintshft("Shift"))  ''Add Trim By  08-11-08
        .TextMatrix(bytTmp, 1) = Format(adRsintshft("Shf_In"), "0.00")
        .TextMatrix(bytTmp, 2) = Format(adRsintshft("Shf_Out"), "0.00")
        .TextMatrix(bytTmp, 3) = adRsintshft("Night")
    End With
    adRsintshft.MoveNext
    bytTmp = bytTmp + 1
Loop
With MSF1
    .TextMatrix(bytTmp, 0) = pVStar.WosCode
    .TextMatrix(bytTmp + 1, 0) = pVStar.HlsCode
End With
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub MSF1_Click()
MSF1.Col = 0
If MSF1.Rows = 1 Then Exit Sub

If Right(strAutoG, 2) = "Em" Then
    If MSF1.Text = "ALL" Then
        txtShifts.Text = "ALL"
    Else
        If Left(txtShifts.Text, 3) = "ALL" Then txtShifts.Text = ""
        txtShifts.Text = txtShifts.Text & MSF1.Text & "."
    End If
Else
    txtShifts.Text = txtShifts.Text & Trim(MSF1.Text) & "."     '' Add trim
End If
''
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then Call MSF1_Click
End Sub
