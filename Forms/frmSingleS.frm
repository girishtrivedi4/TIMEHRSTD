VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSingleS 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3270
      TabIndex        =   1
      Top             =   1680
      Width           =   1005
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3270
      TabIndex        =   2
      Top             =   2190
      Width           =   1005
   End
   Begin VB.TextBox txtShift 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   3330
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   525
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
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
Attribute VB_Name = "frmSingleS"
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
If Trim(txtShift.Text) = "" Then
    MsgBox NewCaptionTxt("50005", adrsC), vbExclamation
    Exit Sub
End If
strDjFileN = txtShift.Text
bytShfMode = 9
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me, True)            '' Sets the Form Icon
Call RetCaptions                '' Set the Captions
Call FillGrid                   '' Fill the ShiftGrid
strDjFileN = ""                 '' Initialize the String to Blank
End Sub

Private Sub FillGrid()
On Error GoTo ERR_P
Dim bytTmp As Byte
If adRsInstall.State = 1 Then adRsInstall.Close
adRsInstall.Open "Select Shift,Shf_In,Shf_Out,Night from InstShft where shift <> '100' Order by Shift" _
, ConMain, adOpenStatic
If Not (adRsInstall.EOF And adRsInstall.BOF) Then
    MSF1.Rows = adRsInstall.RecordCount + 3
    For bytTmp = 1 To adRsInstall.RecordCount
        MSF1.TextMatrix(bytTmp, 0) = adRsInstall("Shift")
        MSF1.TextMatrix(bytTmp, 1) = Format(adRsInstall("Shf_In"), "0.00")
        MSF1.TextMatrix(bytTmp, 2) = Format(adRsInstall("Shf_Out"), "0.00")
        MSF1.TextMatrix(bytTmp, 3) = IIf(adRsInstall("Night") = 0, "N", "Y")
        adRsInstall.MoveNext
    Next
    MSF1.TextMatrix(bytTmp, 0) = pVStar.WosCode
    bytTmp = bytTmp + 1
    MSF1.TextMatrix(bytTmp, 0) = pVStar.HlsCode
Else
    MSF1.Rows = 1
End If
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub RetCaptions()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where ID like '50%'", ConMain, adOpenStatic, adLockReadOnly
Me.Caption = NewCaptionTxt("50001", adrsC)     '' Form Caption
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)
cmdCan.Caption = NewCaptionTxt("00003", adrsMod)
'' Grid Operations
'' Sizing
With MSF1
    .ColWidth(0) = .ColWidth(0) * 0.7
    .ColWidth(1) = .ColWidth(1) * 0.65
    .ColWidth(2) = .ColWidth(2) * 0.65
    .ColWidth(3) = .ColWidth(3) * 0.64
End With
'' Naming
With MSF1
    .TextMatrix(0, 0) = NewCaptionTxt("00031", adrsMod)
    .TextMatrix(0, 1) = NewCaptionTxt("50002", adrsC)
    .TextMatrix(0, 2) = NewCaptionTxt("50003", adrsC)
    .TextMatrix(0, 3) = NewCaptionTxt("50004", adrsC)
End With
With MSF1
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
End With
Exit Sub
ERR_P:
    ShowError ("RetCaptions :: " & Me.Caption)
    Resume Next
End Sub

Private Sub PutShiftValue()
If MSF1.Rows = 1 Then Exit Sub
MSF1.Col = 0
If MSF1.Text = NewCaptionTxt("00031", adrsMod) Then Exit Sub
txtShift.Text = MSF1.Text
End Sub

Private Sub MSF1_Click()
Call PutShiftValue
End Sub

Private Sub MSF1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then Call PutShiftValue
End Sub
