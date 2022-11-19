VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImpLeave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Leaves"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstS 
      Height          =   2400
      Left            =   690
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox lstD 
      Height          =   2400
      Left            =   2010
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   435
      Left            =   780
      TabIndex        =   13
      Top             =   1890
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   435
      Left            =   2190
      TabIndex        =   12
      Top             =   1890
      Width           =   1125
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   435
      Left            =   3810
      TabIndex        =   11
      Top             =   1890
      Width           =   1125
   End
   Begin TabDlg.SSTab sst1 
      Height          =   1875
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3307
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDBC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdFT"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDBC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "optDBC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtFT"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "optFT"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboYear"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdInfo"
      Tab(1).Control(1)=   "cmdBal"
      Tab(1).Control(2)=   "cmdTrn"
      Tab(1).Control(3)=   "cmdDesc"
      Tab(1).ControlCount=   4
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   825
      End
      Begin VB.CommandButton cmdDesc 
         Caption         =   "Leave &Master"
         Height          =   375
         Left            =   -74910
         TabIndex        =   7
         Top             =   330
         Width           =   4815
      End
      Begin VB.CommandButton cmdTrn 
         Caption         =   "Current Leave &Transaction"
         Height          =   375
         Left            =   -74910
         TabIndex        =   10
         Top             =   1410
         Width           =   4815
      End
      Begin VB.CommandButton cmdBal 
         Caption         =   "Current Leave B&alance"
         Height          =   375
         Left            =   -74910
         TabIndex        =   9
         Top             =   1050
         Width           =   4815
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Current Leave &Information"
         Height          =   375
         Left            =   -74910
         TabIndex        =   8
         Top             =   690
         Width           =   4815
      End
      Begin VB.OptionButton optFT 
         Caption         =   "&Free Table Directory"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   2325
      End
      Begin VB.TextBox txtFT 
         Height          =   345
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   870
         Width           =   4095
      End
      Begin VB.OptionButton optDBC 
         Caption         =   ".&DBC"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   1260
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.TextBox txtDBC 
         Height          =   345
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1470
         Width           =   4095
      End
      Begin VB.CommandButton cmdFT 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4350
         TabIndex        =   2
         Top             =   810
         Width           =   525
      End
      Begin VB.CommandButton cmdDBC 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4350
         TabIndex        =   1
         Top             =   1500
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   2460
         TabIndex        =   16
         Top             =   420
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmImpLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' Intrinsic Data Types
Dim bytCtr As Byte, bytLst As Byte
Dim strDjConn As String, strDBPath As String
Dim strBal As String, strInfo As String, strTrn As String
'' Database Variables
Private cnFX As New ADODB.Connection
Private adFX As New ADODB.Recordset
Private adrsImp As New ADODB.Recordset
'' The Database System Objects Variables
Private adCa As New ADOX.Catalog
Private adTa As New ADOX.Table

Private Sub cboYear_Click()
strBal = "lvbal" & Right(cboYear.Text, 2)
strInfo = "lvinfo" & Right(cboYear.Text, 2)
strTrn = "lvtrn" & Right(cboYear.Text, 2)
End Sub

Private Sub cmdBal_Click()
On Error GoTo ERR_Particular
If MsgBox("Are You Sure to Import " & strBal, vbYesNo + vbQuestion) = vbYes Then
    '' VstarDataEnv.cnDJConn.Execute "truncate table " & strBal
    Call TruncateTable(strBal)
    ImportSub (strBal)
End If
Exit Sub
ERR_Particular:
    ShowError ("Balance :: " & Me.Caption)
End Sub

Private Sub cmdDesc_Click()
On Error GoTo ERR_Particular
If MsgBox("Are You Sure to Import Leaves Master File", vbYesNo + vbQuestion) = vbYes Then
    Call ImpLeavDesc
End If
Exit Sub
ERR_Particular:
    ShowError ("Master :: " & Me.Caption)
End Sub

Private Sub cmdInfo_Click()
On Error GoTo ERR_Particular
If MsgBox("Are You Sure to Import " & strInfo, vbYesNo + vbQuestion) = vbYes Then
    ImportSub (strInfo)
End If
Exit Sub
ERR_Particular:
    ShowError ("Information :: " & Me.Caption)
End Sub

Private Sub cmdTrn_Click()
On Error GoTo ERR_Particular
If MsgBox("Are You Sure to Import " & strTrn, vbYesNo + vbQuestion) = vbYes Then
    ImportSub (strTrn)
End If
ERR_Particular:
    ShowError ("Transaction :: " & Me.Caption)
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
Dim intTmp As Integer
For intTmp = 1970 To 2070
    cboYear.AddItem CStr(intTmp)
Next
cboYear.Text = pVStar.YearSel
strBal = "lvbal" & Right(cboYear.Text, 2)
strInfo = "lvinfo" & Right(cboYear.Text, 2)
strTrn = "lvtrn" & Right(cboYear.Text, 2)
''Get the Year
''Get the Yearly Processing File Names
Call RetCaption
sst1.TabEnabled(1) = False
cmdBack.Enabled = False
cmdFT.Enabled = False
txtFT.Enabled = False
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub optFT_Click()
txtDBC.Enabled = False
cmdDBC.Enabled = False
txtFT.Enabled = True
cmdFT.Enabled = True
End Sub

Private Sub optDBC_Click()
txtFT.Enabled = False
cmdFT.Enabled = False
txtDBC.Enabled = True
cmdDBC.Enabled = True
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    blnL = False
End Sub

Private Sub cmdFT_Click()
blnDType = False
strDatPattern = "*.DBF"
frmDBFC.Show vbModal
txtFT.Text = strDjFileN
strDjConn = "dsn=Visual FoxPro Tables;sourcedb=" & strDjFileN & ";sourcetype=dbf;"
End Sub

Private Sub cmdDBC_Click()
blnDType = True
frmDBFC.Show vbModal
txtDBC.Text = strDjFileN
strDjConn = "dsn=Visual FoxPro Database;sourcedb=" & strDjFileN & ";sourcetype=dbc;"
End Sub

Private Sub cmdBack_Click()
sst1.TabEnabled(1) = False
sst1.TabEnabled(0) = True
sst1.Tab = 0
cmdBack.Enabled = False
cmdNext.Enabled = True
End Sub

Private Sub cmdNext_Click()
If Trim(txtDBC.Text) = "" And Trim(txtFT.Text) = "" Then Exit Sub
    Call ConnFX
    If Not blnL Then
        MsgBox Err.Description, vbCritical, App.EXEName
    Else
        sst1.TabEnabled(0) = False
        sst1.TabEnabled(1) = True
        sst1.Tab = 1
        cmdNext.Enabled = False
        cmdBack.Enabled = True
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub ConnFX()
On Error GoTo ERR_Particular
If cnFX.State = 1 Then cnFX.Close
cnFX.Open strDjConn
adCa.ActiveConnection = cnFX
blnL = True
Exit Sub
ERR_Particular:
blnL = False
End Sub

Private Sub ImportSub(ByVal strtab As String)
On Error GoTo ERR_Particular
Dim byttmp As Byte
If Not findT(strtab) Then
        MsgBox "The Table '" & strtab & "' is not found in the requested Data Source", vbCritical, App.EXEName & _
        "  :: Import Data"
        Exit Sub
End If
If CheckFlds(strtab) > 0 Then
        MsgBox "Data Cannot be Imported due to the Inconsistency in Table Columns :: Try Re-Creating Yearly Leave Files", vbCritical, App.EXEName & "  :: Import Data"
        adrsImp.Close
        adFX.Close
        Exit Sub
Else

        If Not adFX.EOF Then adFX.MoveFirst
        Do While Not adFX.EOF
                '' Add New Blank Record
                adrsImp.AddNew
                ''Put Values in the Tables on Field by Field basis
                ''For bytTmp = 0 To adrsImp.Fields.Count - 1
                        ''adrsImp.Fields(CInt(bytTmp)) = adFX.Fields(CInt(bytTmp))
                ''Next
                For byttmp = 0 To lstD.ListCount - 1
                    If Not FindFieldInSource((lstD.List(byttmp))) Then GoTo Rec_Label
                    If adrsImp.Fields(lstD.List(byttmp)).Type = adBoolean Then
                        adrsImp(lstD.List(byttmp)) = IIf(adFX(lstD.List(byttmp)) = "Y", 1, 0)
                    ElseIf adFX.Fields(lstD.List(byttmp)).Type = adDate Or adFX.Fields(lstD.List(byttmp)).Type = adDBDate Then
                        adrsImp(lstD.List(byttmp)) = IIf(adFX(lstD.List(byttmp)) = "12:00:00 AM", Null, _
                        DateCompDate(adFX(lstD.List(byttmp))))
                    Else
                        adrsImp(lstD.List(byttmp)) = Trim(adFX(lstD.List(byttmp)))
                    End If
Rec_Label:
                Next
                adrsImp.Update
                adFX.MoveNext
        Loop
        If adrsImp.State = 1 Then adrsImp.Close
        MsgBox "Database Table  '" & strtab & "' Imported Succesfully", vbInformation, App.EXEName & "  :: Import Data"
End If
Exit Sub
ERR_Particular:
    ShowError ("Import Procedure :: " & Me.Caption)
End Sub

Private Function findT(ByVal strtab As String) As Boolean
On Error GoTo ERR_P
adCa.Tables.Refresh
findT = False
For Each adTa In adCa.Tables
    If UCase(adTa.Name) = UCase(strtab) Then
        findT = True
        Exit For
    End If
Next
Exit Function
ERR_P:
    ShowError ("FindT :: " & Me.Caption)
End Function

Private Function CheckFlds(ByVal strtab As String) As Byte
On Error GoTo ERR_Particular
Dim bytErr As Byte, byttmp As Byte
bytErr = 0
'' Check if no of Fields Match
If adrsImp.State = 1 Then adrsImp.Close
adrsImp.Open "Select * from " & strtab, VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic ''Destination Database
If adFX.State = 1 Then adFX.Close
adFX.Open "select * from " & strtab, cnFX, adOpenStatic                                              '' Source Database
''If adrsImp.Fields.Count <> adFX.Fields.Count Then
        ''bytErr = bytErr + 1
''Else
        '' Start Check if The Fields Line Up Matches
            '' Fill The Destination List Box
            lstD.Clear
            For byttmp = 0 To adrsImp.Fields.Count - 1
                lstD.AddItem UCase(adrsImp.Fields(CInt(byttmp)).Name)
            Next
            ''Fill the Source List Box
            lstS.Clear
            For byttmp = 0 To adFX.Fields.Count - 1
                lstS.AddItem UCase(adFX.Fields(CInt(byttmp)).Name)
            Next
            ''Compare
            ''For byttmp = 0 To lstD.ListCount - 1
                ''If lstD.List(byttmp) <> lstD.List(byttmp) Then bytErr = bytErr + 1
            ''Next
        ''For bytTmp = 0 To adrsImp.Fields.Count - 1
                ''If UCase(adrsImp.Fields(CInt(bytTmp)).Name) <> UCase(adFX.Fields(CInt(bytTmp)).Name) Then _
                ''bytErr = bytErr + 1
        ''Next
        '' End Check if The Fields Line Up Matches
''End If
CheckFlds = bytErr
Exit Function
ERR_Particular:
    ShowError ("Check Fields :: " & Me.Caption)
End Function

Private Sub UpLv()
On Error GoTo ERR_Particular
'' Update pvstar variables for HL,WO,A,P
If adrsLeave.State = 1 Then adrsLeave.Close
adrsLeave.Open "select * from leavdesc where leave='Present Days' or leave='Absent Days' or leave='Weekly Off' or" & _
" leave ='Holiday Days' order by Leave", VstarDataEnv.cnDJConn
If Not (adrsLeave.EOF And adrsLeave.BOF) Then
    adrsLeave.MoveFirst
    pVStar.AbsCode = adrsLeave!lvcode
    adrsLeave.MoveNext
    pVStar.HlsCode = adrsLeave!lvcode
    adrsLeave.MoveNext
    pVStar.PrsCode = adrsLeave!lvcode
    adrsLeave.MoveNext
    pVStar.WosCode = adrsLeave!lvcode
End If
Exit Sub
ERR_Particular:
    ShowError ("UpLv :: " & Me.Caption)
End Sub

Private Sub ImpLeavDesc()
On Error GoTo ERR_Particular
Set adrsImp = Nothing
Set adrsImp = New ADODB.Recordset
VstarDataEnv.cnDJConn.Execute "Delete from leavdesc"
If adrsImp.State = 1 Then adrsImp.Close
adrsImp.Open "Select * from leavdesc", VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic  ''Destination Database
If adFX.State = 1 Then adFX.Close
adFX.Open "select * from leavdesc", cnFX, adOpenStatic   '' Source Database
If adFX.EOF And adFX.BOF Then Exit Sub
Do While Not adFX.EOF
    adrsImp.AddNew
    Select Case UCase(Trim(adFX("Leave")))
        Case "PRESENT"
            If optFT.Value = True Then
                adrsImp("leave") = "Present days"
                adrsImp("isitleave") = "N"
            Else
                adrsImp("isitleave") = "Y"
            End If
        Case "ABSENT"
            If optFT.Value = True Then
                adrsImp("leave") = "Absent Days"
                adrsImp("isitleave") = "N"
            Else
                adrsImp("isitleave") = "Y"
            End If
        Case "PAID_HL"
            If optFT.Value = True Then
                adrsImp("leave") = "Holiday days"
                adrsImp("isitleave") = "N"
            Else
                adrsImp("isitleave") = "Y"
            End If
        Case "WEEKOFF"
            If optFT.Value = True Then
                adrsImp("leave") = "Holiday days"
                adrsImp("isitleave") = "N"
            Else
                adrsImp("isitleave") = "Y"
            End If
        Case Else
            adrsImp("leave") = IIf(IsNull(adFX("leave")), "", adFX("leave"))
            adrsImp("isitleave") = "Y"
    End Select
    adrsImp("lvcode") = adFX("lvcode")
    adrsImp("type") = IIf(IsNull(adFX("type")), "N", adFX("type"))
    adrsImp("Paid") = IIf(IsNull(adFX("Paid")), "N", adFX("Paid"))
    adrsImp("encase") = IIf(IsNull(adFX("encase")), "N", adFX("encase"))
    adrsImp("lv_cof") = IIf(IsNull(adFX("lv_cof")), "N", adFX("lv_cof"))
    adrsImp("lv_qty") = IIf(IsNull(adFX("lv_qty")), 0, adFX("lv_qty"))
    adrsImp("lv_acumul") = IIf(IsNull(adFX("lv_acumul")), 0, adFX("lv_acumul"))
    adrsImp("run_wrk") = IIf(IsNull(adFX("run_wrk")), "N", adFX("run_wrk"))
    ''adrsImp("isitleave") = IIf(IsNull(adFX("isitleave")), "N", adFX("isitleave"))
    adrsImp("cat") = IIf(IsNull(adFX("cat")), "", adFX("cat"))
    adrsImp("creditnow") = IIf(IsNull(adFX("creditnow")), "N", adFX("creditnow"))
    adrsImp("fulcredit") = IIf(IsNull(adFX("fulcredit")), "N", adFX("fulcredit"))
    adrsImp("no_oftimes") = IIf(IsNull(adFX("no_oftimes")), 0, adFX("no_oftimes"))
    adrsImp("allowdays") = IIf(IsNull(adFX("allowdays")), 0, adFX("allowdays"))
    adrsImp("minallowdays") = IIf(IsNull(adFX("minallowda")), 0, adFX("minallowda"))
    adrsImp("custcode") = IIf(IsNull(adFX("custcode")), "", adFX("custcode"))
    adrsImp.Update
    adFX.MoveNext
Loop
MsgBox "Leave Master Imported Successfully", vbInformation, App.EXEName
Call UpLv
Exit Sub
ERR_Particular:
    If Err.Number = 3265 Then Resume Next
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P
sst1.TabCaption(1) = CaptionTxt(46132)   ' Import
optFT.Caption = CaptionTxt(46133)        ' &Free Table Directory
optDBC.Caption = CaptionTxt(46134)       ' .&DBC
cmdBack.Caption = CaptionTxt(46135)      ' &Back
cmdNext.Caption = CaptionTxt(46136)      ' &Next
cmdExit.Caption = CaptionTxt(46137)      ' E&xit
sst1.TabCaption(0) = CaptionTxt(46131)   ' Data Source Info
cmdInfo.Caption = CaptionTxt(46141)      ' Current Leave &Information
cmdTrn.Caption = CaptionTxt(46142)       ' Current Leave &Transaction
cmdBal.Caption = CaptionTxt(46143)       ' Current Leave B&alance
cmdDesc.Caption = CaptionTxt(46144)      ' Leave &Master
Exit Sub
ERR_P:
    ShowError ("RetCaption :: " & Me.Caption)
End Sub

Private Function FindFieldInSource(ByVal strFieldName As String) As Boolean
On Error GoTo ERR_P     '' Checks if the Specific Field Exists in the Source
FindFieldInSource = False
Dim byttmp As Byte
For byttmp = 0 To lstS.ListCount - 1
    If UCase(strFieldName) = UCase(lstS.List(byttmp)) Then FindFieldInSource = True
Next
Exit Function
ERR_P:
    FindFieldInSource = False
End Function
