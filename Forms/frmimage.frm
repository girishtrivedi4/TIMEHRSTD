VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmimage 
   Caption         =   "Image"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2940
   LinkTopic       =   "Form2"
   ScaleHeight     =   1260
   ScaleWidth      =   2940
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtimg 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog ComDilog 
      Left            =   1320
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdbrowse_Click()              ' BY
On Error GoTo Err
    ComDilog.Filter = "Pictures Files (*.bmp)|*.bmp"
    ComDilog.ShowOpen
    If ComDilog.FileName = "" Then Exit Sub
    txtimg.Text = ComDilog.FileName
    'picEmp.Picture = LoadPicture(txtimg.Text)
    cmdsave.Enabled = True
Err:
    If Err.Number = 481 Then
       MsgBox "Invalid Photo.....", vbExclamation, "Error Loading....."
       txtimg.Text = ""
    End If
End Sub

Private Sub cmdsave_Click()
   ConMain.Execute "UPDATE company SET IMG='" & txtimg.Text & "'"
End Sub


