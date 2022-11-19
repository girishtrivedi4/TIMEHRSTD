VERSION 5.00
Begin VB.Form FRMI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Dat File"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   4770
      TabIndex        =   6
      Top             =   270
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4800
      TabIndex        =   7
      Top             =   990
      Width           =   1215
   End
   Begin VB.DriveListBox drv1 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   270
      Width           =   2685
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   2070
      TabIndex        =   3
      Top             =   870
      Width           =   2715
   End
   Begin VB.FileListBox flb1 
      Height          =   3405
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   270
      Width           =   2025
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Drives"
      Height          =   195
      Left            =   1980
      TabIndex        =   0
      Top             =   30
      Width           =   450
   End
   Begin VB.Label lblFolders 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fold&ers"
      Height          =   195
      Left            =   1980
      TabIndex        =   2
      Top             =   630
      Width           =   510
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Files"
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   315
   End
End
Attribute VB_Name = "FRMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adrsC As New ADODB.Recordset

Private Sub cmdCan_Click()
    strDjFileN = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ERR_P
Dim strArrTmp() As String
Dim strDriv As String, intCnt As Integer
Dim bytPos1 As Byte, bytPos2 As Byte
strDjFileN = ""
For intCnt = 0 To flb1.ListCount - 1
    If flb1.Selected(intCnt) = True Then
        If Len(Dir1.List(Dir1.ListIndex)) > 3 Then
            strDjFileN = strDjFileN & "|" & Dir1.List(Dir1.ListIndex) & "\" & flb1.List(intCnt)
        Else
            strDjFileN = strDjFileN & "|" & Dir1.List(Dir1.ListIndex) & flb1.List(intCnt)
        End If
    End If
Next
Unload Me
Exit Sub
ERR_P:
    ShowError ("Select Dat :: " & Me.Caption)
End Sub

Private Sub Dir1_Click()
On Error GoTo Err_particular
    flb1.Path = Dir1.List(Dir1.ListIndex)
Exit Sub
Err_particular:
    MsgBox Err.Description, vbCritical, App.EXEName
End Sub

Private Sub drv1_Change()
On Error GoTo Err_particular
    Dir1.Path = drv1.Drive
    flb1.Path = Dir1.List(Dir1.ListIndex)
Exit Sub
Err_particular:
    MsgBox Err.Description, vbCritical, App.EXEName
    If Err.Number = 68 Then
        drv1.Drive = "C:\"
    End If
End Sub

Private Sub Form_Load()
    flb1.Pattern = "*.Dat"
    Call RetCaption
    Call GotoDataPath
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * from NewCaptions where  ID like '46%'", VstarDataEnv.cnDJConn, adOpenStatic, adLockReadOnly
Call SetFormIcon(Me)                    '' Sets the Form Icon
Me.Caption = NewCaptionTxt("46001", adrsC)          '' Select Dat File
lblFiles.Caption = NewCaptionTxt("46002", adrsC)     '' Files
lblFolders.Caption = NewCaptionTxt("46003", adrsC)   '' Folders
lblDrive.Caption = NewCaptionTxt("46004", adrsC)     '' Drives
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)        '' OK
cmdCan.Caption = NewCaptionTxt("00003", adrsMod)       '' Cancel
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Sub GotoDataPath()              '' Sets the Path to the Path Specified
On Error GoTo ERR_P
Dim strTmp As String
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "Select DatPath from Install", VstarDataEnv.cnDJConn
If Not (adrsLeave.EOF And adrsLeave.BOF) Then
    If (Not IsNull(adrsLeave("DatPath")) And Not IsEmpty(adrsLeave("DatPath"))) Then
        'this if condition add by
        If Not adrsLeave("DatPath") = "" And Not adrsLeave("DatPath") = "" And Not flb1.Path = adrsLeave("DatPath") = "" Then
            drv1.Drive = adrsLeave("DatPath")
            Dir1.Path = adrsLeave("DatPath")
            flb1.Path = adrsLeave("DatPath")
        End If
        If flb1.ListCount > 0 Then flb1.ListIndex = 0
    End If
End If
Exit Sub
ERR_P:
    Select Case Err.Number
        Case 380            '' Path not Found Error
            ShowError ("Specified DAT File Path not Found")
        Case Else
            ShowError ("GotoDataPath :: " & Me.Caption)
    End Select
End Sub
