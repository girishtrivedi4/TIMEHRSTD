VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportUtil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Data"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   435
      Left            =   3840
      TabIndex        =   11
      Top             =   2250
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   435
      Left            =   2190
      TabIndex        =   7
      Top             =   2250
      Width           =   1125
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   435
      Left            =   780
      TabIndex        =   10
      Top             =   2250
      Width           =   1125
   End
   Begin TabDlg.SSTab sst1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3942
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "optFT"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtFT"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optDBC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDBC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdFT"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDBC"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdImp"
      Tab(1).Control(1)=   "lstTables"
      Tab(1).ControlCount=   2
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
         TabIndex        =   6
         Top             =   1620
         Width           =   525
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
         TabIndex        =   3
         Top             =   870
         Width           =   525
      End
      Begin VB.TextBox txtDBC 
         Height          =   345
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1620
         Width           =   4095
      End
      Begin VB.OptionButton optDBC 
         Caption         =   ".&DBC"
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   1320
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.TextBox txtFT 
         Height          =   345
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   870
         Width           =   4095
      End
      Begin VB.OptionButton optFT 
         Caption         =   "&Free Table Directory"
         Height          =   225
         Left            =   210
         TabIndex        =   1
         Top             =   540
         Width           =   2325
      End
      Begin VB.ListBox lstTables 
         Height          =   1815
         Left            =   -74850
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2115
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Import"
         Height          =   585
         Left            =   -72630
         TabIndex        =   9
         Top             =   1590
         Width           =   2475
      End
   End
   Begin VB.ListBox lstD 
      Height          =   2010
      Left            =   1470
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   30
      Width           =   1335
   End
   Begin VB.ListBox lstS 
      Height          =   2010
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   30
      Width           =   1335
   End
End
Attribute VB_Name = "frmImportUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' Intrinsic Data Types
Dim bytCtr As Byte, bytLst As Byte
Dim strTabName(12) As String, strSelTab() As String, strValTab() As String
Dim strInv As String, strErrMsg As String
Dim strDjConn As String, strDBPath As String
'' Database Variables
Private cnFX As New ADODB.Connection
Private adFX As New ADODB.Recordset
Private adrsImp As New ADODB.Recordset
'' The Database System Objects Variables
Private adCa As New ADOX.Catalog
Private adTa As New ADOX.Table

Private Sub cmdBack_Click()
sst1.TabEnabled(1) = False
sst1.TabEnabled(0) = True
sst1.Tab = 0
cmdBack.Enabled = False
cmdNext.Enabled = True
End Sub

Private Sub cmdDBC_Click()
blnDType = True
frmDBFC.Show vbModal
txtDBC.Text = strDjFileN
strDjConn = "dsn=Visual FoxPro Database;sourcedb=" & strDjFileN & ";sourcetype=dbc;"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFT_Click()
blnDType = False
strDatPattern = "*.DBF"
frmDBFC.Show vbModal
txtFT.Text = strDjFileN
strDjConn = "dsn=Visual FoxPro Tables;sourcedb=" & strDjFileN & ";sourcetype=dbf;"
End Sub

Private Sub cmdImp_Click()      'Import
On Error GoTo ERR_Particular
Dim bytctr1 As Byte
bytctr1 = 0
'' Check if any of items to be Imported is Selected or not
For bytCtr = 0 To bytLst
        If lstTables.Selected(bytCtr) = True Then bytctr1 = bytctr1 + 1
Next
If bytctr1 = 0 Then
        MsgBox "Please Select the Table to be Imported", vbExclamation, App.EXEName & "  :: Import"
        Exit Sub
End If
'' Make Array of the Selected List Items
Call MakeSelArr
Dim strMsg As String
For bytCtr = 0 To UBound(strSelTab)
        strMsg = strMsg & vbCrLf & vbTab & CStr(bytCtr + 1) & ". " & strSelTab(bytCtr)
Next
'' Confirm the Import
If MsgBox("Are you sure to Import the folowing Database Tables ?" & strMsg, vbYesNo _
+ vbQuestion, App.EXEName & " :: Import Data") = vbYes Then
    For bytCtr = 0 To UBound(strSelTab)
        Select Case UCase(strSelTab(bytCtr))
            ''Case "CATDESC"
                ''Call ImportCatDesc
            Case Else
                Call ImportSub(strSelTab(bytCtr))
        End Select
    Next
End If
Exit Sub
ERR_Particular:
    ShowError ("Import :: " & Me.Caption)
End Sub

Private Sub cmdNext_Click()
If Trim(txtDBC.Text) = "" And Trim(txtFT.Text) = "" Then Exit Sub
    Call ConnFX
    If Not blnL Then
        ShowError ("Next :: " & Me.Caption)
    Else
        sst1.TabEnabled(0) = False
        sst1.TabEnabled(1) = True
        sst1.Tab = 1
        cmdNext.Enabled = False
        cmdBack.Enabled = True
    End If
End Sub

Private Sub cmdSel_Click()      'Select All Items
For bytCtr = 0 To bytLst
        lstTables.Selected(bytCtr) = True
Next
End Sub

Private Sub cmdUnsel_Click()    'Unselect All Items
For bytCtr = 0 To bytLst
    lstTables.Selected(bytCtr) = False
Next
End Sub

Private Sub Form_Load()
On Error GoTo ERR_P
bytLst = lstTables.ListCount - 1
''Put the Names of the Master filesin the strTabName Array
For bytCtr = 0 To bytLst
    strTabName(bytCtr) = Choose(bytCtr + 1, "catdesc", _
     "declwohl", "deptdesc", "empmst", "groupmst", _
    "holiday", "lost", "ro_shift", "instshft")
Next
'' Check if the Connection is Eastablished
Call RetCaption
sst1.TabEnabled(1) = False
cmdBack.Enabled = False
cmdFT.Enabled = False
txtFT.Enabled = False
Exit Sub
ERR_P:
    ShowError ("Load :: " & Me.Caption)
End Sub

Private Sub MakeSelArr()        'Procedure to make Array of the Selected Items
On Error GoTo ERR_P
Dim intCtr As Integer
intCtr = -1
For bytCtr = 0 To bytLst
        If lstTables.Selected(bytCtr) = True Then
                ReDim Preserve strSelTab(intCtr + 1)
                strSelTab(UBound(strSelTab)) = strTabName(bytCtr)
                intCtr = intCtr + 1
        End If
Next
Exit Sub
ERR_P:
    ShowError ("MakeSelArr :: " & Me.Caption)
End Sub

Private Function RetValidTabs() As Byte
On Error GoTo ERR_Particular
strInv = ""
Dim bytRet As Byte, bytNo As Byte
bytRet = 0: bytNo = 1
For bytCtr = 0 To UBound(strSelTab)
        If adrsImp.State = 1 Then adrsImp.Close
        adrsImp.Open "select * from " & strSelTab(bytCtr)
        If Not (adrsImp.EOF And adrsImp.BOF) Then
                strInv = strInv & vbCrLf & vbTab & vbTab & CStr(bytNo) & ".  " & strSelTab(bytCtr)
                bytNo = bytNo + 1
        Else
                ReDim Preserve strValTab(bytRet)
                strValTab(bytRet) = strSelTab(bytCtr)
                bytRet = bytRet + 1
        End If
Next
RetValidTabs = bytRet
Exit Function
ERR_Particular:
    ShowError ("Return Valid Tables :: " & Me.Caption)
End Function

Private Sub ImportSub(ByVal strtab As String)
On Error GoTo ERR_Particular
Set adrsImp = Nothing
Set adrsImp = New ADODB.Recordset
Dim byttmp As Byte
If Not findT(strtab) Then
        MsgBox "The Table '" & strtab & "' is not found in the requested Data Source", vbCritical, App.EXEName & _
        "  :: Import Data"
        Exit Sub

Else
    If UCase(strtab) = "INSTSHFT" Then Call TruncateTable("INSTSHFT")
End If
If CheckFlds(strtab) > 0 Then       '' Put Fields in Sorting Order
Else
        If Not adFX.EOF Then adFX.MoveFirst
        Do While Not adFX.EOF
                '' Add New Blank Record
                adrsImp.AddNew
                For byttmp = 0 To lstD.ListCount - 1
                    '' If Field is not Available in the Data Source
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
                If Not adFX.EOF Then        '' If Data is from Clipper Adjust for Same CAT
                    If UCase(strtab) = "CATDESC" Then
                        If UCase(adFX("CAT")) = UCase(adrsImp("Cat")) Then adFX.MoveNext
                    End If
                End If
        Loop
        If optFT.Value = True Then      '' If the Data is from Clipper Adjust Missing Values
                If UCase(strtab) = "EMPMST" Then
                    '' Group
                    VstarDataEnv.cnDJConn.Execute "Update Empmst Set Group=0"
                    '' Company
                    Select Case inVar.bytCom
                        Case "1", "", "0"
                            VstarDataEnv.cnDJConn.Execute "Update Empmst Set Company=1"
                    End Select
                    '' JoinDate
                    VstarDataEnv.cnDJConn.Execute "Update Empmst set JoinDate=" & strDTEnc & _
                    DateCompStr(CStr(Date)) & strDTEnc & " Where JoinDate is NULL"
                    '' Shift Date
                    VstarDataEnv.cnDJConn.Execute "Update Empmst set Shf_date=" & strDTEnc & _
                    DateCompStr(CStr(Date)) & strDTEnc & " Where Shf_date is NULL"
                End If
        End If
        If adrsImp.State = 1 Then adrsImp.Close
        MsgBox "Database Table  '" & strtab & "' Imported Succesfully", vbInformation, App.EXEName & "  :: Import Data"
End If
Exit Sub
ERR_Particular:
    Select Case Err.Number
        Case 3265           '' Item Not Found
          Err.Clear
          GoTo Rec_Label
        Case -2147217900    '' Duplicate Values (Primary Key)
            adrsImp.Cancel
            Resume Next
        Case 3219           '' Operation is Not Allowed in this Context
            Set adrsImp = Nothing
            Set adrsImp = New ADODB.Recordset
            Resume Next
        Case Else           '' Other Error
            If UCase(strtab) = "EMPMST" Then
                Resume Next '' For Multi-Step Operation Generating Error
            Else            '' Very Unknown Error
                ShowError ("Import Procedure :: " & Me.Caption & lstD.List(byttmp))
            End If
    End Select
End Sub

Private Function CheckFlds(ByVal strtab As String) As Byte
On Error GoTo ERR_Particular
Dim bytErr As Byte, byttmp As Byte
bytErr = 0
'' Check if no of Fields Match
If adrsImp.State = 1 Then adrsImp.Close
adrsImp.Open "Select * from " & strtab, VstarDataEnv.cnDJConn, adOpenKeyset, adLockOptimistic ''Destination Database
If adFX.State = 1 Then adFX.Close
adFX.Open "select * from " & strtab, cnFX, adOpenStatic                                              '' Source Database
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
CheckFlds = bytErr
Exit Function
ERR_Particular:
    Select Case Err.Number
        Case -2147217865
            If UCase(strtab) = "RO_SHIFT" Then strtab = "SCODE": Resume
        Case Else
            ShowError ("Check Fields :: " & Me.Caption)
    End Select
End Function

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

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_P
'' Set the Global Variables Free
''Intrinsic data Types
Erase strTabName
Erase strSelTab
Erase strValTab
'' Database Objects
Set cnFX = Nothing
Set adFX = Nothing
'' Database System Objects
Set adCa = Nothing
Set adTa = Nothing
Exit Sub
ERR_P:
    ShowError ("Ok :: " & Me.Caption)
End Sub

Private Function findT(ByVal strtab As String) As Boolean
On Error GoTo ERR_P
ABCD:
adCa.Tables.Refresh
findT = False
For Each adTa In adCa.Tables
        If UCase(adTa.Name) = UCase(strtab) Then
                findT = True
                Exit For
        End If
Next
If findT = False And UCase(strtab) = "RO_SHIFT" Then
    strtab = "SCODE"
    GoTo ABCD
End If
Exit Function
ERR_P:
    If findT = False And UCase(strtab) = "RO_SHIFT" Then
        strtab = "SCODE"
        GoTo ABCD
    End If
    ShowError ("FindT :: " & Me.Caption)
End Function

Private Sub optDBC_Click()
txtFT.Enabled = False
cmdFT.Enabled = False
txtDBC.Enabled = True
cmdDBC.Enabled = True
End Sub

Private Sub optFT_Click()
txtDBC.Enabled = False
cmdDBC.Enabled = False
txtFT.Enabled = True
cmdFT.Enabled = True
End Sub

Private Function FileCheck(ByVal strFileN As String) As Boolean
FileCheck = True
If Trim(strFileN) = "" Then FileCheck = False
End Function

Private Sub sst1_Click(PreviousTab As Integer)
blnL = False
End Sub

Private Sub RetCaption()
On Error GoTo ERR_P
sst1.TabCaption(1) = CaptionTxt(46132)   '' Import
optFT.Caption = CaptionTxt(46133)        '' &Free Table Directory
optDBC.Caption = CaptionTxt(46134)       '' .&DBC
cmdBack.Caption = CaptionTxt(46135)      '' &Back
cmdNext.Caption = CaptionTxt(46136)      '' &Next
cmdExit.Caption = CaptionTxt(46137)      '' E&xit
cmdImp.Caption = CaptionTxt(46140)       '' &Import
sst1.TabCaption(0) = CaptionTxt(46131)   '' Data Source Info
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
