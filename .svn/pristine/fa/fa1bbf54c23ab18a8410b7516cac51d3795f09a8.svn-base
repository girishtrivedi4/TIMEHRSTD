VERSION 5.00
Begin VB.Form frmEntries 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Dat File"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDat 
      Caption         =   "Select Dat File"
      Height          =   3825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2220
         TabIndex        =   4
         Top             =   3420
         Width           =   1065
      End
      Begin VB.CommandButton cmdCan 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3300
         TabIndex        =   5
         Top             =   3420
         Width           =   1095
      End
      Begin VB.ListBox lstDat 
         Height          =   3180
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   4335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   3420
         Width           =   1065
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1140
         TabIndex        =   3
         Top             =   3420
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''
Dim adrsC As New ADODB.Recordset

Private Sub cmdAdd_Click()
On Error GoTo ERR_P
Dim strTmp() As String, intCnt As Integer
frmSelectDat.Show vbModal
If strDjFileN <> "" Then
    strTmp = Split(strDjFileN, "|")
    For intCnt = LBound(strTmp) To UBound(strTmp)
        If Trim(strTmp(intCnt)) <> "" Then lstDat.AddItem strTmp(intCnt)
    Next
    ''lstDat.AddItem strDjFileN
End If
Exit Sub
ERR_P:
    ShowError ("Error Adding .DAT Files :: " & Me.Caption)
    Resume Next
End Sub

Private Sub cmdCan_Click()
    frmReports.optDly(0).Value = True
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
    MsgBox NewCaptionTxt("00062", adrsC), vbExclamation
    Exit Sub
Else
    If lstDat.ListCount = 0 Then
        MsgBox NewCaptionTxt("00050", adrsC), vbExclamation
        cmdAdd.SetFocus
        Exit Sub
    End If
    adrsRits("Daily") = 1
    adrsRits.Update
    adrsRits.Close
End If
'' End
If lstDat.ListCount = 0 Then
    MsgBox NewCaptionTxt("00050", adrsC), vbExclamation
    cmdAdd.SetFocus
Else
    '' Code for Daily Process
    If Not AppendDataFile(lstDat, Me) Then
        MsgBox NewCaptionTxt("17011", adrsC), vbCritical, App.EXEName
        Exit Sub
    End If
    typDT.dtFrom = DateCompDate(typRep.strDlyDate)
    typDT.dtTo = DateCompDate(typRep.strDlyDate)
    Call FillInstalltypes       '' Fills Details from Parameters to their respective Type
    Call OpenMasters            '' Opens Necessary Master Tables
    '' Phase 4
    Call FillEntEmpFound        '' Fill EmpFound Table With All the Employees
    Call FilterEmpty            '' Clears all the Records from tbldata which are blank
    Call FilterOnDates          '' Filter on the basis of Dates
    Call TruncateTable("DailyPro")
    Call GetDataPunches         '' Puts Data in processing Table
    Call FilterOnCard           '' Filter on the basis of Cards
    Call PutEmpCode             '' Puts Employee Code to all Records
    Call PutFlag                '' Puts Flag to all Records
    If adrsLeave.State = 1 Then adrsLeave.Close
    adrsLeave.Open "Select * from dailypro", VstarDataEnv.cnDJConn, adOpenKeyset
    If Not (adrsLeave.EOF And adrsLeave.BOF) Then
        MsgBox NewCaptionTxt("26003", adrsC) & typRep.strDlyDate, vbInformation
    Else
        MsgBox NewCaptionTxt("26004", adrsC) & typRep.strDlyDate, vbInformation
        frmReports.optDly(0).Value = True
    End If
    Call SetDailyFlag        '' Reset Flag
    Unload Me
End If
Exit Sub
ERR_P:
    ShowError ("Entries :: " & Me.Caption)
End Sub

Private Sub FillEntEmpFound()
On Error GoTo ERR_P
Call TruncateTable("EmpfoundTb")
VstarDataEnv.cnDJConn.Execute "Insert into EmpFoundtb(EmpCode) Select Distinct EmpCode from Empmst"
Exit Sub
ERR_P:
    ShowError ("FillEntEmpFound :: " & Me.Caption)
End Sub

Private Sub Form_Load()
Call SetFormIcon(Me, True)        '' Load Form Icon
Call RetCaptions            '' Set Form Captions
End Sub

Private Sub RetCaptions()
'On Error Resume Next
If adrsC.State = 1 Then adrsC.Close
adrsC.Open "Select * From NewCaptions Where ID Like '17%' or ID Like '26%' or ID Like '00%'", VstarDataEnv.cnDJConn, adOpenStatic
Me.Caption = NewCaptionTxt("26001", adrsC)      '' Form Caption
cmdAdd.Caption = NewCaptionTxt("00004", adrsC)  '' Add
cmdDel.Caption = NewCaptionTxt("26002", adrsC)  '' Remove
cmdOK.Caption = NewCaptionTxt("00002", adrsC)   '' OK
cmdCan.Caption = NewCaptionTxt("00003", adrsC)  '' Cancel
frDat.Caption = NewCaptionTxt("26001", adrsC)   '' Select Dat File
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then Call ShowF10("26")
If KeyCode = vbKeyF10 Then KeyCode = 0
End Sub

