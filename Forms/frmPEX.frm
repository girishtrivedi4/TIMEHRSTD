VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDecrypt 
      Height          =   435
      Left            =   6900
      TabIndex        =   6
      Top             =   3450
      Width           =   1365
   End
   Begin VB.CommandButton cmdEncrypt 
      Height          =   435
      Left            =   5550
      TabIndex        =   5
      Top             =   3450
      Width           =   1365
   End
   Begin VB.Frame frDetails 
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtU 
         Height          =   345
         Left            =   2340
         TabIndex        =   4
         Top             =   3030
         Width           =   5925
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2505
         Left            =   2340
         TabIndex        =   3
         Top             =   510
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   4419
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.TextBox txtQ 
         Height          =   315
         Left            =   2340
         TabIndex        =   2
         Top             =   180
         Width           =   5925
      End
      Begin VB.ListBox lstTables 
         Height          =   3180
         Left            =   30
         TabIndex        =   1
         Top             =   180
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' PE Module
Option Explicit

Private Sub cmdDecrypt_Click()
    Call InputBox("Decryted String is", "", DEncryptDat(InputBox("Encrypt String", App.EXEName, "", _
    ScaleWidth / 2, ScaleHeight / 2), 1))
End Sub

Private Sub cmdEncrypt_Click()
    Call InputBox("Encrypted String is", "", DEncryptDat(InputBox("Encrypt String", App.EXEName, "", _
    ScaleWidth / 2, ScaleHeight / 2), 1))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Call SetFormIcon(Me)
Call AddTables
Exit Sub
ERR_P:
    ShowError ("PE :: " & Me.Caption)
End Sub

Private Sub lstTables_DblClick()
If lstTables.ListCount <= 0 Then Exit Sub
If lstTables.ListIndex < 0 Then Exit Sub
txtQ.Text = "Select * from " & lstTables.List(lstTables.ListIndex)
Call DoSelect
End Sub

Private Sub lstTables_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call lstTables_DblClick
End Sub

Private Sub AddTables()         '' Add User Tables to the List
On Error GoTo ERR_P
Dim adTmp As New ADODB.Recordset, bytTmp As Byte
adTmp.CursorType = adOpenStatic
adTmp.LockType = adLockOptimistic
Set adTmp = VstarDataEnv.cnDJConn.OpenSchema(adSchemaTables)
lstTables.Clear
If adTmp.RecordCount = 0 Then Exit Sub
For bytTmp = 0 To adTmp.RecordCount - 1
    If adTmp(3) = "TABLE" Then lstTables.AddItem UCase(adTmp(2))
    adTmp.MoveNext
Next
Exit Sub
ERR_P:
    ShowError (" Add Tables :: " & Me.Caption)
End Sub

Private Sub txtQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call DoSelect
End Sub

Private Sub DoSelect()          '' Process  the Select Query
If Len(Trim(txtQ.Text)) < 10 Then Exit Sub
If InStr(UCase(Trim(txtQ.Text)), "SELECT") <= 0 Then
    MsgBox "Please Write the Select Query", vbExclamation, App.EXEName
    Exit Sub
End If
Select Case ExeSel
    Case 0      '' Error
        MSF1.Rows = 1
        MSF1.Cols = 1
        Me.Caption = ""
    Case 1      '' Succesfull
        MSF1.Rows = 1
        MSF1.Cols = 1
        Call FillGrid
    Case 2      '' No Records
End Select
End Sub

Private Function ExeSel() As Byte       '' Execute Select Statement
On Error GoTo ERR_P
ExeSel = 1
If adrsTemp.State = 1 Then adrsTemp.Close
adrsTemp.Open Trim(txtQ.Text), VstarDataEnv.cnDJConn, adOpenKeyset
If (adrsTemp.EOF And adrsTemp.BOF) Then
    MsgBox "No Records Found", vbExclamation, App.EXEName
    ExeSel = 2
End If
Exit Function
ERR_P:
    ShowError ("Execute Select :: " & Me.Caption)
    ExeSel = 0
End Function

Private Sub FillGrid()                  '' Fill the Grid with the Execute Statement
On Error GoTo ERR_P
Dim bytTmp As Byte, intTmp As Integer
'' Get All Column Names
MSF1.Cols = adrsTemp.Fields.Count
For bytTmp = 0 To adrsTemp.Fields.Count - 1
    MSF1.TextMatrix(0, bytTmp) = UCase(adrsTemp.Fields(CInt(bytTmp)).Name)
Next
'' Get All Rows
MSF1.Rows = adrsTemp.RecordCount + 1
adrsTemp.MoveFirst
For intTmp = 1 To adrsTemp.RecordCount
    For bytTmp = 0 To adrsTemp.Fields.Count - 1
        MSF1.TextMatrix(intTmp, bytTmp) = IIf(IsNull(adrsTemp.Fields(CInt(bytTmp)).Value), "", _
        adrsTemp.Fields(CInt(bytTmp)).Value)
    Next
    adrsTemp.MoveNext
Next
Me.Caption = adrsTemp.RecordCount
Exit Sub
ERR_P:
    ShowError ("FillGrid :: " & Me.Caption)
End Sub

Private Sub DoUpdate()                      '' Process the Update Query
If Len(Trim(txtU.Text)) < 10 Then Exit Sub
''If InStr(UCase(Trim(txtU.Text)), "UPDATE") <= 0 Then
''    Msgbox "Please Write the Update Query", vbExclamation, App.EXEName
''    Exit Sub
''End If
If MsgBox("Are You Sure to Execute this Query", vbYesNo + vbQuestion, App.EXEName) = vbYes Then
Else
    Exit Sub
End If
Select Case ExeUpd
    Case 0      '' Error
    Case 1      '' Succesfull
        MsgBox "Operation Successfull", vbExclamation, App.EXEName
    Case 2      '' Other Consequences
End Select
End Sub

Private Function ExeUpd() As Byte       '' Execute Update Statement
On Error GoTo ERR_P
ExeUpd = 1
VstarDataEnv.cnDJConn.Execute Trim(txtU.Text)
Exit Function
ERR_P:
    ShowError ("Execute Update :: " & Me.Caption)
    ExeUpd = 0
End Function

Private Sub txtU_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call DoUpdate
End Sub
