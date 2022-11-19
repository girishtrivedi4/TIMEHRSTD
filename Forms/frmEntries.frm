VERSION 5.00
Begin VB.Form frmEntries 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Dat File"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDat 
      Caption         =   "Select Dat File"
      Height          =   3825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5205
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2610
         TabIndex        =   4
         Top             =   3420
         Width           =   1215
      End
      Begin VB.CommandButton cmdCan 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3870
         TabIndex        =   5
         Top             =   3420
         Width           =   1245
      End
      Begin VB.ListBox lstDat 
         Height          =   3180
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   5085
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   3420
         Width           =   1215
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1350
         TabIndex        =   3
         Top             =   3420
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSelEmp  As String
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdAdd_Click()
On Error GoTo ERR_P
Dim strTmp() As String, intCnt As Integer

If strDjFileN <> "" Then
    strTmp = Split(strDjFileN, "|")
    For intCnt = LBound(strTmp) To UBound(strTmp)
        If Trim(strTmp(intCnt)) <> "" Then lstDat.AddItem strTmp(intCnt)
    Next
End If
Exit Sub
ERR_P:
    ShowError ("Error Adding .DAT Files :: " & Me.Caption)
    'Resume Next
End Sub

Private Sub cmdCan_Click()
'    frmReports.optDly(0).Value = True
    Unload Me
End Sub

Private Sub cmdDel_Click()
If lstDat.ListIndex <> -1 Then
    lstDat.RemoveItem lstDat.ListIndex
End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo ERR_P
'' Check if Any Lock is there
If adrsRits.State = 1 Then adrsRits.Close
adrsRits.Open "Select * from Exc", VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic
If adrsRits("Daily") = 1 Then
    adrsRits.Close
    MsgBox NewCaptionTxt("00062", adrsMod), vbExclamation
    Exit Sub
Else
    If lstDat.ListCount = 0 Then
        MsgBox NewCaptionTxt("00050", adrsMod), vbExclamation
        cmdAdd.SetFocus
        Exit Sub
    End If
    adrsRits("Daily") = 1
    adrsRits.Update
    adrsRits.Close
End If
'' End
If lstDat.ListCount = 0 Then
    MsgBox NewCaptionTxt("00050", adrsMod), vbExclamation
    cmdAdd.SetFocus
Else
    '' Code for Daily Process
    If Not AppendDataFile(Me) Then
        MsgBox NewCaptionTxt("17011", adrsC), vbCritical, App.EXEName
        Exit Sub
    End If
       typDT.dtFrom = DateCompDate(typRep.strDlyDate)
       typDT.dtTo = DateCompDate(typRep.strDlyDate)
 
    Call FillInstalltypes       '' Fills Details from Parameters to their respective Type
    Call MakestrSelEmp
    Call OpenMasters(strSelEmp)            '' Opens Necessary Master Tables
    '' Phase 4
    Call FilterEmpty            '' Clears all the Records from tbldata which are blank
    Call FilterOnDates          '' Filter on the basis of Dates
    Call TruncateTable("DailyPro")
    
    If (typPerm.blnIO Or typPerm.blnDI) Then
        Call GetDataPunchesIO(strSelEmp)           '' Puts Data in processing Table
    Else
        Call GetDataPunches(strSelEmp)           '' Puts Data in processing Table
    End If
    ''
    Call FilterOnCard           '' Filter on the basis of Cards
    Call PutFlag                '' Puts Flag to all Records
    Call FilterOnTime
    If adrsLeave.State = 1 Then adrsLeave.Close
    adrsLeave.Open "Select * from dailypro", VstarDataEnv.cnDJConn, adOpenKeyset
        If Not (adrsLeave.EOF And adrsLeave.BOF) Then
            MsgBox NewCaptionTxt("26003", adrsC) & typRep.strDlyDate, vbInformation
        Else
            MsgBox NewCaptionTxt("26004", adrsC) & typRep.strDlyDate, vbInformation
        '        frmReports.optDly(0).Value = True
        End If

    Call SetDailyFlag        '' Reset Flag
    Unload Me
End If
Exit Sub
ERR_P:
    ShowError ("Entries :: " & Me.Caption)
End Sub

Private Sub MakestrSelEmp()
On Error GoTo ERR_P
Dim adrsDumb As New ADODB.Recordset
strSelEmp = ""
If adrsDumb.State = 1 Then adrsDumb.Close
adrsDumb.ActiveConnection = VstarDataEnv.cnDJConn
adrsDumb.CursorType = adOpenStatic
adrsDumb.LockType = adLockReadOnly
adrsDumb.Open "Select Empmst.Empcode from " & rpTables & _
" Where Empmst.Empcode=Empmst.Empcode  " & strSql
Do While Not adrsDumb.EOF
    strSelEmp = strSelEmp & "'" & adrsDumb("Empcode") & "',"
    adrsDumb.MoveNext
Loop
If strSelEmp <> "" Then
    strSelEmp = Left(strSelEmp, Len(strSelEmp) - 1)
    strSelEmp = "(" & strSelEmp & ")"
End If
Exit Sub
ERR_P:
    ShowError ("MakestrSelEmp :: " & Me.Caption)
    '
    Resume Next
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me, True)        '' Load Form Icon
Call RetCaptions            '' Set Form Captions
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '17%' or ID Like '26%'", VstarDataEnv.cnDJConn, adOpenStatic
Me.Caption = NewCaptionTxt("26001", adrsC)      '' Form Caption
cmdAdd.Caption = NewCaptionTxt("00004", adrsMod)  '' Add
cmdDel.Caption = NewCaptionTxt("26002", adrsC)  '' Remove
cmdOk.Caption = NewCaptionTxt("00002", adrsMod)   '' OK
cmdCan.Caption = NewCaptionTxt("00003", adrsMod)  '' Cancel
frDat.Caption = NewCaptionTxt("26001", adrsC)   '' Select Dat File
End Sub
