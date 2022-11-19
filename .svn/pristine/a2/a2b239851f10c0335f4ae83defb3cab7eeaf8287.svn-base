Attribute VB_Name = "mdlMain"
Option Explicit
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias _
   "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
   cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type DSNDetails
    BackEnd As Byte
    DSNName As String
    Database As String
    UserName As String
    Password As String
    ServerName As String
    Path As String
End Type
Public TDSN As DSNDetails
Public VstarConn As New ADODB.Connection

Public Function CreateTABLES() As Boolean
Select Case TDSN.BackEnd
Case 1 ''SQL - Server
    Call CreateDatabase
    VstarConn.Execute "Use VStarDB"
    VstarConn.Execute "Use VStarDB"
    VstarConn.Execute "Use VStarDB"
Case 3 ''Oracle
    Call CreateSequence
End Select
If Not FindTable("TBLINDEXMASTER") Then CreateTBLINDEXMASTER
If Not FindTable("CATDESC") Then CreateCATDESC
If Not FindTable("COMPANY") Then CreateCOMPANY
If Not FindTable("CORUL") Then CreateCORUL
If Not FindTable("DAILYPRO") Then CreateDAILYPRO
If Not FindTable("DECLWOHL") Then CreateDECLWOHL
If Not FindTable("DEPTDESC") Then CreateDEPTDESC
If Not FindTable("DIVISION") Then CreateDIVISION
If Not FindTable("DPERF") Then CreateDPERF
If Not FindTable("DPRAB") Then CreateDPRAB
If Not FindTable("DSUM") Then CreateDSUM
If Not FindTable("DTRN") Then CreateDTRN
If Not FindTable("EMPFOUNDTB") Then CreateEMPFOUNDTB
If Not FindTable("EXC") Then CreateEXC
If Not FindTable("GROUPMST") Then CreateGROUPMST
If Not FindTable("LOCATION") Then CreateLOCATION
If Not FindTable("HOLIDAY") Then CreateHOLIDAY
If Not FindTable("INSTALL") Then CreateINSTALL
If Not FindTable("INSTSHFT") Then CreateINSTSHFT
If Not FindTable("LATEERL") Then CreateLATEERL
If Not FindTable("LOST") Then CreateLOST
If Not FindTable("LVT") Then CreateLVT
If Not FindTable("LVTRNPERMT") Then CreateLVTRNPERMT
If Not FindTable("MALE") Then CreateMALE
If Not FindTable("MATT") Then CreateMATT
If Not FindTable("MEMOTABLE") Then CreateMEMOTABLE
If Not FindTable("MONTRN") Then CreateMONTRN
If Not FindTable("MPERF") Then CreateMPERF
If Not FindTable("NEWCAPTIONS") Then CreateNEWCAPTIONS
If Not FindTable("OTRUL") Then CreateOTRUL
If Not FindTable("PPERF") Then CreatePPERF
If Not FindTable("RO_SHIFT") Then CreateRO_SHIFT
If Not FindTable("SHFINFO") Then CreateSHFINFO
If Not FindTable("EMPMST") Then CreateEMPMST
If Not FindTable("LEAVBAL") Then CreateLEAVBAL
If Not FindTable("LEAVDESC") Then CreateLEAVDESC
If Not FindTable("LEAVINFO") Then CreateLEAVINFO
If Not FindTable("LEAVTRN") Then CreateLEAVTRN
If Not FindTable("TBLDATA") Then CreateTBLDATA
If Not FindTable("USERACCS") Then CreateUSERACCS
If Not FindTable("WPERF") Then createWPerf
If Not FindTable("WSTAT") Then CreateWSTAT
If Not FindTable("YRTB") Then CreateYRTB
If Not FindTable("Dsumc") Then createtblDsumc
If Not FindTable("MonperfA") Then createMonPerfA
If Not FindTable("Monperfb") Then createMonperfB
If Not FindTable("yrAbPr") Then createyrAbPr
If Not FindTable("errorLog") Then createerrorLog

CreateTABLES = True
End Function

Public Sub DropTables()
On Error GoTo Err_P
Select Case TDSN.BackEnd
Case 1 ''SQL - Server
   '' Do Nothing
Case 3 ''Oracle
    Call DeleteConn("Next1", 1)
End Select
Call DeleteConn("LEAVBAL")
Call DeleteConn("HOLIDAY")
Call DeleteConn("LEAVDESC")
Call DeleteConn("EMPMST")
Call DeleteConn("CATDESC")
Call DeleteConn("COMPANY")
Call DeleteConn("CORUL")
Call DeleteConn("DAILYPRO")
Call DeleteConn("DECLWOHL")
Call DeleteConn("DEPTDESC")
Call DeleteConn("DIVISION")
Call DeleteConn("DPERF")
Call DeleteConn("DPRAB")
Call DeleteConn("DSUM")
Call DeleteConn("DTRN")
Call DeleteConn("EMPFOUNDTB")
Call DeleteConn("EXC")
Call DeleteConn("GROUPMST")
Call DeleteConn("INSTALL")
Call DeleteConn("INSTSHFT")
Call DeleteConn("LATEERL")
Call DeleteConn("LEAVINFO")
Call DeleteConn("LEAVTRN")
Call DeleteConn("LOCATION")
Call DeleteConn("LOST")
Call DeleteConn("LVT")
Call DeleteConn("LVTRNPERMT")
Call DeleteConn("MALE")
Call DeleteConn("MATT")
Call DeleteConn("MEMOTABLE")
Call DeleteConn("MONTRN")
Call DeleteConn("MPERF")
Call DeleteConn("NEWCAPTIONS")
Call DeleteConn("OTRUL")
Call DeleteConn("PPERF")
Call DeleteConn("RO_SHIFT")
Call DeleteConn("SHFINFO")
Call DeleteConn("TBLDATA")
Call DeleteConn("USERACCS")
Call DeleteConn("WPERF")
Call DeleteConn("WSTAT")
Call DeleteConn("YRTB")
Call DeleteConn("MONPERFA")
Call DeleteConn("MONPERFB")
Call DeleteConn("DSUMC")
Call DeleteConn("YRABPR")
Call DeleteConn("ErrorLog")
Call DeleteConn("TBLINDEXMASTER")
Exit Sub
Err_P:
    ShowError ("DropTables::mdlMain")
End Sub

Private Sub DeleteConn(strTName As String, Optional bytSeq As Byte = 0)
On Error GoTo Err_P
Dim StrTmp As String
frmCreate.lblDisp.Caption = "Deleting => " & strTName
frmCreate.lblDisp.Refresh
If bytSeq = 0 Then
    StrTmp = "Drop Table " & strTName
Else
    StrTmp = "Drop Sequence " & strTName
End If
VstarConn.Execute StrTmp
frmCreate.lblDisp.Caption = ""
frmCreate.lblDisp.Refresh
Exit Sub
Err_P:
    If Err.Number = -2147467259 Or Err.Number = -2147217865 Then
        ''Already Exists then continue
    Else
        StrTmp = "There is some problem Deleting Table/Sequence " & strTName & vbCrLf & _
        "Do you wish to continue?" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
        "Please check for the following ..." & vbCrLf & vbTab & _
        "-> Database connectivity is no longer available." & vbCrLf & vbTab & "-> The User '" & _
        TDSN.UserName & "' does not have enough previleges for Table/Sequence Deletion."
        If MsgBox(StrTmp, vbYesNo + vbQuestion) = vbNo Then
            End
        End If
    End If
End Sub


Private Sub CreateDatabase()
On Error GoTo Err_P
VstarConn.Execute "Use Master"
VstarConn.Execute "Use Master"
VstarConn.Execute "Use Master"
VstarConn.Execute "Use Master"
VstarConn.Execute "Create Database VStarDB"
Sleep (5000)
DoEvents
DoEvents
DoEvents
Sleep (5000)
DoEvents
DoEvents
VstarConn.Execute "EXEC sp_dboption 'vstardb', 'into/bulkcopy', 'TRUE'"
VstarConn.Execute "Use VStarDB"
VstarConn.Execute "Use VStarDB"
VstarConn.Execute "Use VStarDB"
VstarConn.Execute "Use VStarDB"
Sleep (5000)
Exit Sub
Err_P:
    If Err.Number = -2147217900 Then
        ''Database already exists
    Else
        MsgBox Err.Description
    End If
End Sub

Private Function FindTable(ByVal strName As String) As Boolean
On Error GoTo ERR_Particular
Dim adrsFT As New ADODB.Recordset
Select Case TDSN.BackEnd
    Case 1  '' SQL Server
        adrsFT.Open "Select count(*) from SysObjects where Name='" & strName & "'" _
        , VstarConn
    Case 3  ''Oracle
        adrsFT.Open "select count(*) from Tabs where UPPER(Table_Name) ='" & _
        UCase(strName) & "'", _
        VstarConn
End Select
If IsEmpty(adrsFT(0)) Or IsNull(adrsFT(0)) Or adrsFT(0) <= 0 Then
    FindTable = False
Else
    FindTable = True
End If
adrsFT.Close
Exit Function
ERR_Particular:
    ShowError ("FindTable::" & strName)
End Function

Public Sub ShowError(strmsg As String)
MsgBox Err.Description & ":::" & strmsg
End Sub

Public Sub createtblDsumc()
 Dim StrTmp As String
 Select Case TDSN.BackEnd
 
 Case 1 '' Sql-Server
 StrTmp = " CREATE TABLE Dsumc (Empcode char (10)  NULL ,EmpName nvarchar (50)  NULL ," & _
          " cat nvarchar (10)  NULL ,catdescdesc nvarchar (50)  NULL ,dept smallint NULL ," & _
          " deptdescdesc nvarchar (50)  NULL ,Location smallint NULL ,LocDesc nvarchar (50)  NULL ," & _
          " Div smallint NULL ,Divdesc nvarchar (50)  NULL ,[Group] nvarchar (10)  NULL ,grupdesc nvarchar (50)  NULL ," & _
          " Company tinyint NULL ,Cname nvarchar (50)  NULL ,sec nvarchar (15)  NULL ,SecDesc nvarchar (50)  NULL ," & _
          " BaseC nvarchar (50)  NULL ,Strength int NULL ,NofE int NULL ,PP real NULL ,PPP real NULL ,AA real NULL ," & _
          " AAP real NULL ,WO real NULL ,WOP real NULL ,HL real NULL ,HLP real NULL ,PL real NULL ,PLP real NULL ," & _
          " UPL real NULL ,UPLP real NULL ,OD real NULL ,ODP real NULL ,OTHRS real NULL ,TOT real NULL ," & _
          " TOTP real NULL ,OT real NULL ) ON PRIMARY GO"

Case 3 ''Oracle
 
  StrTmp = "CREATE TABLE Dsumc (Empcode varchar2 (10)NULL , EmpName varchar2 (50)  NULL ," & _
           " cat varchar2 (10)  NULL , catdescdesc varchar2 (50)  NULL , dept smallint NULL ," & _
           " deptdescdesc varchar2 (50)  NULL , Location smallint NULL , LocDesc varchar2 (50)  NULL ," & _
           " Div smallint NULL , Divdesc varchar2 (50)  NULL , ""Group"" varchar2 (10)  NULL ," & _
           " grupdesc varchar2 (50)  NULL , Company smallint NULL , Cname varchar2 (50)  NULL ," & _
           " sec varchar2 (15)  NULL , SecDesc varchar2 (50)  NULL , BaseC varchar2 (50)  NULL ," & _
           " Strength int NULL , NofE int NULL , PP number NULL , PPP number NULL , AA number NULL ," & _
           " AAP number NULL , WO number NULL , WOP number NULL , HL number NULL , HLP number NULL ," & _
           " PL number NULL ,PLP number NULL ,UPL number NULL ,UPLP number NULL ,OD number NULL ,ODP number NULL ," & _
           " OTHRS number NULL ,TOT number NULL ,TOTP number NULL ,OT number NULL)"
 
 
End Select
Call ExecuteConn(StrTmp, "DsumC")
End Sub

Public Sub createMonPerfA()
Dim StrTmp As String
 
 Select Case TDSN.BackEnd
   Case 1 '' Sql-Server
       
    StrTmp = "CREATE TABLE MonperfA (Empcode varchar (10)  NULL ,Arr1 float NULL ,Arr2 float NULL ,Arr3 float NULL ,Arr4 float NULL ," & _
             " Arr5 float NULL ,Arr6 float NULL ,Arr7 float NULL ,Arr8 float NULL ,Arr9 float NULL ,Arr10 float NULL ,Arr11 float NULL ," & _
             " Arr12 float NULL ,Arr13 float NULL ,Arr14 float NULL ,Arr15 float NULL ,Arr16 float NULL ,Arr17 float NULL ," & _
             " Arr18 float NULL ,Arr19 float NULL ,Arr20 float NULL ,Arr21 float NULL ,Arr22 float NULL ,Arr23 float NULL ," & _
             " Arr24 float NULL ,Arr25 float NULL ,Arr26 float NULL ,Arr27 float NULL ,Arr28 float NULL ,Arr29 float NULL ," & _
             " Arr30 float NULL ,Arr31 float NULL ,Dep1 float NULL ,Dep2 float NULL ,Dep3 float NULL ,Dep4 float NULL ,Dep5 float NULL ," & _
             " Dep6 float NULL ,Dep7 float NULL ,Dep8 float NULL ,Dep9 float NULL ,Dep10 float NULL ,Dep11 float NULL ,Dep12 float NULL ," & _
             " Dep13 float NULL ,Dep14 float NULL ,Dep15 float NULL ,Dep16 float NULL ,Dep17 float NULL ,Dep18 float NULL ,Dep19 float NULL ," & _
             " Dep20 float NULL ,Dep21 float NULL ,Dep22 float NULL ,Dep23 float NULL ,Dep24 float NULL ," & _
             " Dep25 float NULL ,Dep26 float NULL ,Dep27 float NULL ,Dep28 float NULL ,Dep29 float NULL ,Dep30 float NULL ," & _
             " Dep31 float NULL ,Late1 float NULL ,Late2 float NULL ,Late3 float NULL ,Late4 float NULL ,Late5 float NULL ,Late6 float NULL ,Late7 float NULL ," & _
             " Late8 float NULL ,Late9 float NULL ,Late10 float NULL ,Late11 float NULL ,Late12 float NULL ,Late13 float NULL ," & _
             " Late14 float NULL ,Late15 float NULL ,Late16 float NULL ,Late17 float NULL ,Late18 float NULL ,Late19 float NULL ,Late20 float NULL ," & _
             " Late21 float NULL ,Late22 float NULL ,Late23 float NULL ,Late24 float NULL ,Late25 float NULL ,Late26 float NULL ," & _
             " Late27 float NULL ,Late28 float NULL ,Late29 float NULL ,Late30 float NULL ,Late31 float NULL ,Earl1 float NULL ," & _
             " Earl2 float NULL ,Earl3 float NULL ,Earl4 float NULL ,Earl5 float NULL ,Earl6 float NULL ,Earl7 float NULL ," & _
             " Earl8 float NULL ,Earl9 float NULL ,Earl10 float NULL ,Earl11 float NULL ,Earl12 float NULL ,Earl13 float NULL ," & _
             " Earl14 float NULL ,Earl15 float NULL ,Earl16 float NULL ,Earl17 float NULL ,Earl18 float NULL ," & _
             " Earl19 float NULL ,Earl20 float NULL ,Earl21 float NULL ,Earl22 float NULL ,Earl23 float NULL ,Earl24 float NULL ," & _
             " Earl25 float NULL ,Earl26 float NULL ,Earl27 float NULL ,Earl28 float NULL ,Earl29 float NULL ,Earl30 float NULL ,Earl31 float NULL ," & _
             " Work1 float NULL ,Work2 float NULL ,Work3 float NULL ,Work4 float NULL ,Work5 float NULL ,Work6 float NULL ,Work7 float NULL ,Work8 float NULL ,Work9 float NULL ," & _
             " Work10 float NULL ,Work11 float NULL ,Work12 float NULL ,Work13 float NULL ,Work14 float NULL ,Work15 float NULL ,Work16 float NULL ,Work17 float NULL ,Work18 float NULL ,Work19 float NULL ,Work20 float NULL ,Work21 float NULL ,Work22 float NULL ,Work23 float NULL ,Work24 float NULL ,Work25 float NULL ,Work26 float NULL ,Work27 float NULL ,Work28 float NULL ,Work29 float NULL ,Work30 float NULL ,Work31 float NULL ," & _
             " Dt1 float NULL ,Dt2 float NULL ,Dt3 float NULL ,Dt4 float NULL ,Dt5 float NULL ,Dt6 float NULL ,Dt7 float NULL ,Dt8 float NULL ,Dt9 float NULL ,Dt10 float NULL ,Dt11 float NULL ,Dt12 float NULL ,Dt13 float NULL ,Dt14 float NULL ,Dt15 float NULL ,Dt16 float NULL ,Dt17 float NULL ,Dt18 float NULL ,Dt19 float NULL ,Dt20 float NULL ,Dt21 float NULL ,Dt22 float NULL ,Dt23 float NULL ,Dt24 float NULL ,Dt25 float NULL ,Dt26 float NULL ,Dt27 float NULL ,Dt28 float NULL ,Dt29 float NULL ,Dt30 float NULL ,Dt31 float NULL ," & _
             " OT1 float NULL ,OT2 float NULL ,OT3 float NULL ,OT4 float NULL ,OT5 float NULL ,OT6 float NULL ,OT7 float NULL ,OT8 float NULL ,OT9 float NULL ,OT10 float NULL ,OT11 float NULL ,OT12 float NULL ,OT13 float NULL ,OT14 float NULL ,OT15 float NULL ,OT16 float NULL ,OT17 float NULL ,OT18 float NULL ,OT19 float NULL ,OT20 float NULL ,OT21 float NULL ,OT22 float NULL ,OT23 float NULL ,OT24 float NULL ,OT25 float NULL ,OT26 float NULL ,OT27 float NULL ,OT28 float NULL ,OT29 float NULL ,OT30 float NULL ,OT31 float NULL ) ON PRIMARY GO"
 
    Case 3 '' Oracle
         
         StrTmp = "CREATE TABLE MonperfA (Empcode varchar2 (10)  NULL ,Arr1 number(6,3) NULL ,Arr2 number(6,3) NULL ,Arr3 number(6,3) NULL ,Arr4 number(6,3) NULL ," & _
             " Arr5 number(6,3) NULL ,Arr6 number(6,3) NULL ,Arr7 number(6,3) NULL ,Arr8 number(6,3) NULL ,Arr9 number(6,3) NULL ,Arr10 number(6,3) NULL ,Arr11 number(6,3) NULL ," & _
             " Arr12 number(6,3) NULL ,Arr13 number(6,3) NULL ,Arr14 number(6,3) NULL ,Arr15 number(6,3) NULL ,Arr16 number(6,3) NULL ,Arr17 number(6,3) NULL ," & _
             " Arr18 number(6,3) NULL ,Arr19 number(6,3) NULL ,Arr20 number(6,3) NULL ,Arr21 number(6,3) NULL ,Arr22 number(6,3) NULL ,Arr23 number(6,3) NULL ," & _
             " Arr24 number(6,3) NULL ,Arr25 number(6,3) NULL ,Arr26 number(6,3) NULL ,Arr27 number(6,3) NULL ,Arr28 number(6,3) NULL ,Arr29 number(6,3) NULL ," & _
             " Arr30 number(6,3) NULL ,Arr31 number(6,3) NULL ,Dep1 number(6,3) NULL ,Dep2 number(6,3) NULL ,Dep3 number(6,3) NULL ,Dep4 number(6,3) NULL ,Dep5 number(6,3) NULL ," & _
             " Dep6 number(6,3) NULL ,Dep7 number(6,3) NULL ,Dep8 number(6,3) NULL ,Dep9 number(6,3) NULL ,Dep10 number(6,3) NULL ,Dep11 number(6,3) NULL ,Dep12 number(6,3) NULL ," & _
             " Dep13 number(6,3) NULL ,Dep14 number(6,3) NULL ,Dep15 number(6,3) NULL ,Dep16 number(6,3) NULL ,Dep17 number(6,3) NULL ,Dep18 number(6,3) NULL ,Dep19 number(6,3) NULL ," & _
             " Dep20 number(6,3) NULL ,Dep21 number(6,3) NULL ,Dep22 number(6,3) NULL ,Dep23 number(6,3) NULL ,Dep24 number(6,3) NULL ," & _
             " Dep25 number(6,3) NULL ,Dep26 number(6,3) NULL ,Dep27 number(6,3) NULL ,Dep28 number(6,3) NULL ,Dep29 number(6,3) NULL ,Dep30 number(6,3) NULL ," & _
             " Dep31 number(6,3) NULL ,Late1 number(6,3) NULL ,Late2 number(6,3) NULL ,Late3 number(6,3) NULL ,Late4 number(6,3) NULL ,Late5 number(6,3) NULL ,Late6 number(6,3) NULL ,Late7 number(6,3) NULL ," & _
             " Late8 number(6,3) NULL ,Late9 number(6,3) NULL ,Late10 number(6,3) NULL ,Late11 number(6,3) NULL ,Late12 number(6,3) NULL ,Late13 number(6,3) NULL ," & _
             " Late14 number(6,3) NULL ,Late15 number(6,3) NULL ,Late16 number(6,3) NULL ,Late17 number(6,3) NULL ,Late18 number(6,3) NULL ,Late19 number(6,3) NULL ,Late20 number(6,3) NULL ," & _
             " Late21 number(6,3) NULL ,Late22 number(6,3) NULL ,Late23 number(6,3) NULL ,Late24 number(6,3) NULL ,Late25 number(6,3) NULL ,Late26 number(6,3) NULL ," & _
             " Late27 number(6,3) NULL ,Late28 number(6,3) NULL ,Late29 number(6,3) NULL ,Late30 number(6,3) NULL ,Late31 number(6,3) NULL ,Earl1 number(6,3) NULL ," & _
             " Earl2 number(6,3) NULL ,Earl3 number(6,3) NULL ,Earl4 number(6,3) NULL ,Earl5 number(6,3) NULL ,Earl6 number(6,3) NULL ,Earl7 number(6,3) NULL ," & _
             " Earl8 number(6,3) NULL ,Earl9 number(6,3) NULL ,Earl10 number(6,3) NULL ,Earl11 number(6,3) NULL ,Earl12 number(6,3) NULL ,Earl13 number(6,3) NULL ," & _
             " Earl14 number(6,3) NULL ,Earl15 number(6,3) NULL ,Earl16 number(6,3) NULL ,Earl17 number(6,3) NULL ,Earl18 number(6,3) NULL ," & _
             " Earl19 number(6,3) NULL ,Earl20 number(6,3) NULL ,Earl21 number(6,3) NULL ,Earl22 number(6,3) NULL ,Earl23 number(6,3) NULL ,Earl24 number(6,3) NULL ," & _
             " Earl25 number(6,3) NULL ,Earl26 number(6,3) NULL ,Earl27 number(6,3) NULL ,Earl28 number(6,3) NULL ,Earl29 number(6,3) NULL ,Earl30 number(6,3) NULL ,Earl31 number(6,3) NULL ," & _
             " Work1 number(6,3) NULL ,Work2 number(6,3) NULL ,Work3 number(6,3) NULL ,Work4 number(6,3) NULL ,Work5 number(6,3) NULL ,Work6 number(6,3) NULL ,Work7 number(6,3) NULL ,Work8 number(6,3) NULL ,Work9 number(6,3) NULL ," & _
             " Work10 number(6,3) NULL ,Work11 number(6,3) NULL ,Work12 number(6,3) NULL ,Work13 number(6,3) NULL ,Work14 number(6,3) NULL ,Work15 number(6,3) NULL ,Work16 number(6,3) NULL ,Work17 number(6,3) NULL ,Work18 number(6,3) NULL ,Work19 number(6,3) NULL ,Work20 number(6,3) NULL ,Work21 number(6,3) NULL ,Work22 number(6,3) NULL ,Work23 number(6,3) NULL ,Work24 number(6,3) NULL ,Work25 number(6,3) NULL ,Work26 number(6,3) NULL ,Work27 number(6,3) NULL ,Work28 number(6,3) NULL ,Work29 number(6,3) NULL ,Work30 number(6,3) NULL ,Work31 number(6,3) NULL ," & _
             " Dt1 number(6,3) NULL ,Dt2 number(6,3) NULL ,Dt3 number(6,3) NULL ,Dt4 number(6,3) NULL ,Dt5 number(6,3) NULL ,Dt6 number(6,3) NULL ,Dt7 number(6,3) NULL ,Dt8 number(6,3) NULL ,Dt9 number(6,3) NULL ,Dt10 number(6,3) NULL ,Dt11 number(6,3) NULL ,Dt12 number(6,3) NULL ,Dt13 number(6,3) NULL ,Dt14 number(6,3) NULL ,Dt15 number(6,3) NULL ,Dt16 number(6,3) NULL ,Dt17 number(6,3) NULL ,Dt18 number(6,3) NULL ,Dt19 number(6,3) NULL ,Dt20 number(6,3) NULL ,Dt21 number(6,3) NULL ,Dt22 number(6,3) NULL ,Dt23 number(6,3) NULL ,Dt24 number(6,3) NULL ,Dt25 number(6,3) NULL ,Dt26 number(6,3) NULL ,Dt27 number(6,3) NULL ,Dt28 number(6,3) NULL ,Dt29 number(6,3) NULL ,Dt30 number(6,3) NULL ,Dt31 number(6,3) NULL ," & _
             " OT1 number(6,3) NULL ,OT2 number(6,3) NULL ,OT3 number(6,3) NULL ,OT4 number(6,3) NULL ,OT5 number(6,3) NULL ,OT6 number(6,3) NULL ,OT7 number(6,3) NULL ,OT8 number(6,3) NULL ,OT9 number(6,3) NULL ,OT10 number(6,3) NULL ,OT11 number(6,3) NULL ,OT12 number(6,3) NULL ,OT13 number(6,3) NULL ,OT14 number(6,3) NULL ,OT15 number(6,3) NULL ,OT16 number(6,3) NULL ,OT17 number(6,3) NULL ,OT18 number(6,3) NULL ,OT19 number(6,3) NULL ,OT20 number(6,3) NULL ,OT21 number(6,3) NULL ,OT22 number(6,3) NULL ,OT23 number(6,3) NULL ,OT24 number(6,3) NULL ,OT25 number(6,3) NULL ,OT26 number(6,3) NULL ,OT27 number(6,3) NULL ,OT28 number(6,3) NULL ,OT29 number(6,3) NULL ,OT30 number(6,3) NULL ,OT31 number(6,3) NULL ) "
 
    
   
   End Select

Call ExecuteConn(StrTmp, "MonPerFA")
End Sub

Public Sub createMonperfB()

Dim StrTmp As String

 Select Case TDSN.BackEnd
 
 Case 1 'SQl-Server
 
        StrTmp = "CREATE TABLE MonperfB (Empcode varchar (10)  NULL ,Rem1 nvarchar (5)  NULL ,Rem2 nvarchar (5)  NULL ,Rem3 nvarchar (5)  NULL ,Rem4 nvarchar (5)  NULL ,Rem5 nvarchar (5)  NULL ,Rem6 nvarchar (5)  NULL ,Rem7 nvarchar (5)  NULL ,Rem8 nvarchar (5)  NULL ,Rem9 nvarchar (5)  NULL ," & _
                 " Rem10 nvarchar (5)  NULL ,Rem11 nvarchar (5)  NULL ,Rem12 nvarchar (5)  NULL ,Rem13 nvarchar (5)  NULL ,Rem14 nvarchar (5)  NULL ,Rem15 nvarchar (5)  NULL ,Rem16 nvarchar (5)  NULL ,Rem17 nvarchar (5)  NULL ,Rem18 nvarchar (5)  NULL ,Rem19 nvarchar (5)  NULL ,Rem20 nvarchar (5)  NULL ,Rem21 nvarchar (5)  NULL ," & _
                 " Rem22 nvarchar (5)  NULL ,Rem23 nvarchar (5)  NULL ,Rem24 nvarchar (5)  NULL ,Rem25 nvarchar (5)  NULL ,Rem26 nvarchar (5)  NULL ,Rem27 nvarchar (5)  NULL ,Rem28 nvarchar (5)  NULL ,Rem29 nvarchar (5)  NULL ,Rem30 nvarchar (5)  NULL ,Rem31 nvarchar (5)  NULL ,shf1 nvarchar (5)  NULL ,shf2 nvarchar (5)  NULL ,shf3 nvarchar (5)  NULL ,shf4 nvarchar (5)  NULL ,shf5 nvarchar (5)  NULL ,shf6 nvarchar (5)  NULL ,shf7 nvarchar (5)  NULL ," & _
                 " shf8 nvarchar (5)  NULL ,shf9 nvarchar (5)  NULL ,shf10 nvarchar (5)  NULL ,shf11 nvarchar (5)  NULL ,shf12 nvarchar (5)  NULL ,shf13 nvarchar (5)  NULL ,shf14 nvarchar (5)  NULL ,shf15 nvarchar (5)  NULL ,shf16 nvarchar (5)  NULL ,shf17 nvarchar (5)  NULL ,shf18 nvarchar (5)  NULL ,shf19 nvarchar (5)  NULL ,shf20 nvarchar (5)  NULL ," & _
                 " shf21 nvarchar (5)  NULL ,shf22 nvarchar (5)  NULL ,shf23 nvarchar (5)  NULL ,shf24 nvarchar (5)  NULL ,shf25 nvarchar (5)  NULL ,shf26 nvarchar (5)  NULL ,shf27 nvarchar (5)  NULL ,shf28 nvarchar (5)  NULL ,shf29 nvarchar (5)  NULL ,shf30 nvarchar (5)  NULL ,shf31 nvarchar (5)  NULL) ON PRIMARY GO"
                      
Case 3 ''Oracle
  
      StrTmp = "CREATE TABLE MonperfB (Empcode varchar2(10)  NULL ,Rem1 varchar2 (5)  NULL ,Rem2 varchar2 (5)  NULL ,Rem3 varchar2 (5)  NULL ,Rem4 varchar2 (5)  NULL ,Rem5 varchar2 (5)  NULL ,Rem6 varchar2 (5)  NULL ,Rem7 varchar2 (5)  NULL ,Rem8 varchar2 (5)  NULL ,Rem9 varchar2 (5)  NULL ," & _
                 " Rem10 varchar2 (5)  NULL ,Rem11 varchar2 (5)  NULL ,Rem12 varchar2 (5)  NULL ,Rem13 varchar2 (5)  NULL ,Rem14 varchar2 (5)  NULL ,Rem15 varchar2 (5)  NULL ,Rem16 varchar2 (5)  NULL ,Rem17 varchar2 (5)  NULL ,Rem18 varchar2 (5)  NULL ,Rem19 varchar2 (5)  NULL ,Rem20 varchar2 (5)  NULL ,Rem21 varchar2 (5)  NULL ," & _
                 " Rem22 varchar2 (5)  NULL ,Rem23 varchar2 (5)  NULL ,Rem24 varchar2 (5)  NULL ,Rem25 varchar2 (5)  NULL ,Rem26 varchar2 (5)  NULL ,Rem27 varchar2 (5)  NULL ,Rem28 varchar2 (5)  NULL ,Rem29 varchar2 (5)  NULL ,Rem30 varchar2 (5)  NULL ,Rem31 varchar2 (5)  NULL ,shf1 varchar2 (5)  NULL ,shf2 varchar2 (5)  NULL ,shf3 varchar2 (5)  NULL ,shf4 varchar2 (5)  NULL ,shf5 varchar2 (5)  NULL ,shf6 varchar2 (5)  NULL ,shf7 varchar2 (5)  NULL ," & _
                 " shf8 varchar2 (5)  NULL ,shf9 varchar2 (5)  NULL ,shf10 varchar2 (5)  NULL ,shf11 varchar2 (5)  NULL ,shf12 varchar2 (5)  NULL ,shf13 varchar2 (5)  NULL ,shf14 varchar2 (5)  NULL ,shf15 varchar2 (5)  NULL ,shf16 varchar2 (5)  NULL ,shf17 varchar2 (5)  NULL ,shf18 varchar2 (5)  NULL ,shf19 varchar2 (5)  NULL ,shf20 varchar2 (5)  NULL ," & _
                 " shf21 varchar2 (5)  NULL ,shf22 varchar2 (5)  NULL ,shf23 varchar2 (5)  NULL ,shf24 varchar2 (5)  NULL ,shf25 varchar2 (5)  NULL ,shf26 varchar2 (5)  NULL ,shf27 varchar2 (5)  NULL ,shf28 varchar2 (5)  NULL ,shf29 varchar2 (5)  NULL ,shf30 varchar2 (5)  NULL ,shf31 varchar2 (5)  NULL)"
 End Select
 Call ExecuteConn(StrTmp, "MONPERFB")
End Sub

Public Sub createerrorLog()
Dim StrTmp As String
 
 Select Case TDSN.BackEnd
 
 Case 1 '' Sql-Server]
 StrTmp = "CREATE TABLE errorLog (ErrorName nvarchar (50) NULL ,ErrorId float NULL) ON PRIMARY GO"
 Case 2 ''Oracle
 StrTmp = "CREATE TABLE errorLog (ErrorName varchar2 (50) NULL ,ErrorId number NULL)"
 End Select
 
 Call ExecuteConn(StrTmp, "ErrorLog")
End Sub

Public Sub createyrAbPr()
Dim StrTmp As String
 
 Select Case TDSN.BackEnd
 
 Case 1 '' Sql-Server
   
    StrTmp = "CREATE TABLE yrAbPr (Empcode nvarchar (8)  NOT NULL ,jan float NULL ,feb float NULL ,mar float NULL ," & _
             " apr float NULL ,may float NULL ,jun float NULL ,jul float NULL ,aug float NULL ,sep float NULL ," & _
             " oct float NULL ,nov float NULL ,dec float NULL ,Total float NULL ) ON PRIMARY GO"

  Case 3 '' Oracle
     StrTmp = "CREATE TABLE yrAbPr (Empcode varchar2 (8)  NOT NULL ,jan number NULL ,feb number NULL ,mar number NULL ," & _
              " apr number NULL ,may number NULL ,jun number NULL ,jul number NULL ,aug number NULL ,sep number NULL ," & _
              " oct number NULL ,nov number NULL ,dec number NULL ,Total number NULL)"
 End Select
 
 Call ExecuteConn(StrTmp, "YRABPR")
End Sub
Public Sub createWPerf()

Dim StrTmp As String

Select Case TDSN.BackEnd
 
  
  Case 1 '' SQL-SERVER
        StrTmp = "CREATE TABLE WPerf (Empcode nvarchar (50)  NULL ,d1 nvarchar (8)  NULL ,d2 nvarchar (8)  NULL ,d3 nvarchar (8)  NULL,d4 nvarchar NULL ,d5 nvarchar (8)  NULL ,d6 nvarchar (8)  NULL ," & _
                 " d7 nvarchar (8)  NULL ,arr1 nvarchar (8)  NULL ,arr2 nvarchar (8)  NULL ,arr3 nvarchar (8)  NULL ,arr4 nvarchar (8)  NULL ,arr5 nvarchar (8)  NULL ," & _
                 " arr6 nvarchar (8)  NULL ,arr7 nvarchar (8)  NULL ,dep1 nvarchar (8)  NULL ,dep2 nvarchar (8)  NULL ,dep3 nvarchar (8)  NULL ,dep4 nvarchar (8)  NULL ,dep5 nvarchar (8)  NULL ," & _
                 " dep6 nvarchar (8)  NULL ,dep7 nvarchar (8)  NULL ,late1 float NULL ,late2 float NULL ,late3 float NULL ,late4 float NULL ,late5 float NULL ," & _
                 " late6 float NULL ,late7 float NULL ,erly1 float NULL ,erly2 float NULL ,erly3 float NULL ,erly4 float NULL ,erly5 float NULL ,erly6 float NULL ,erly7 float NULL ,OT1 float NULL ,OT2 float NULL ,OT3 float NULL ,OT4 float NULL ,OT5 float NULL ,OT6 float NULL ,OT7 float NULL ,wrk1 float NULL ,wrk2 float NULL ,wrk3 float NULL ,wrk4 float NULL ,wrk5 float NULL ,wrk6 float NULL ," & _
                 " wrk7 float NULL ,pres1 nvarchar (8)  NULL ,pres2 nvarchar (8)  NULL ,pres3 nvarchar (8)  NULL ,pres4 nvarchar (8)  NULL ,pres5 nvarchar (8)  NULL ,pres6 nvarchar (8)  NULL ,pres7 nvarchar (8)  NULL ,shft1 nvarchar (8)  NULL ,shft2 nvarchar (8)  NULL ,shft3 nvarchar (8)  NULL ,shft4 nvarchar (8)  NULL ,shft5 nvarchar (8)  NULL ,shft6 nvarchar (8)  NULL ,shft7 nvarchar (8)  NULL ,SumExtra float NULL ) ON PRIMARY GO"
     
      
  Case 3 ''ORACLE
  
     StrTmp = "CREATE TABLE WPerf (Empcode varchar2 (50)  NULL ,d1 varchar2 (8)  NULL ,d2 varchar2 (8)  NULL ,d3 varchar2 (8)  NULL ,d4 varchar2 (8)  NULL ,d5 varchar2 (8)  NULL ,d6 varchar2 (8)  NULL ,d7 varchar2 (8)  NULL ," & _
               " arr1 varchar2 (8)  NULL ,arr2 varchar2 (8)  NULL ,arr3 varchar2 (8)  NULL ,arr4 varchar2 (8)  NULL ,arr5 varchar2 (8)  NULL ,arr6 varchar2 (8)  NULL ,arr7 varchar2 (8)  NULL ,dep1 varchar2 (8)  NULL ,dep2 varchar2 (8)  NULL ,dep3 varchar2 (8)  NULL ,dep4 varchar2 (8)  NULL ,dep5 varchar2 (8)  NULL ,dep6 varchar2 (8)  NULL ,dep7 varchar2 (8)  NULL ,late1 number  NULL ,late2 number  NULL ,late3 number  NULL ,late4 number  NULL ,late5 number  NULL ,late6 number  NULL ,late7 number  NULL ,erly1 number  NULL ,erly2 number  NULL ,erly3 number  NULL ,erly4 number  NULL ,erly5 number  NULL ,erly6 number  NULL ,erly7 number  NULL ,OT1 number  NULL ,OT2 number  NULL ,OT3 number  NULL ,OT4 number  NULL ,OT5 number  NULL ,OT6 number  NULL ,OT7 number  NULL ,wrk1 number  NULL ,wrk2 number  NULL ,wrk3 number  NULL ," & _
               " wrk4 number  NULL ,wrk5 number  NULL ,wrk6 number  NULL ,wrk7 number  NULL ,pres1 varchar2 (8)  NULL ,pres2 varchar2 (8)  NULL ,pres3 varchar2 (8)  NULL ,pres4 varchar2 (8)  NULL ," & _
               " pres5 varchar2 (8)  NULL ,pres6 varchar2 (8)  NULL ,pres7 varchar2 (8)  NULL ,shft1 varchar2 (8)  NULL ,shft2 varchar2 (8)  NULL ,shft3 varchar2 (8)  NULL ,shft4 varchar2 (8)  NULL ,shft5 varchar2 (8)  NULL ,shft6 varchar2 (8)  NULL ,shft7 varchar2 (8)  NULL ,SumExtra number NULL )"

  
  Call ExecuteConn(StrTmp, "WPERF")
  End Select
 End Sub

Public Sub CreateTBLINDEXMASTER()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = " CREATE TABLE [dbo].[TBLINDEXMASTER] (" & _
               " CTBLNAME NVARCHAR(25) NOT NULL , CUNIQUE NVARCHAR(10) NULL ," & _
               " CINDEXKEYS NVARCHAR(50) NOT NULL ,  CINDEXNAME NVARCHAR(25) NOT NULL)"

    Case 3 ''oracle
        StrTmp = "Create Table TblIndexMaster( CTBLNAME VARCHAR2(25) NOT NULL ," & _
                " CUNIQUE VARCHAR2(10) NULL , CINDEXKEYS VARCHAR2(50) NOT NULL ," & _
                " CINDEXNAME VARCHAR2(25) NOT NULL)"
End Select
Call ExecuteConn(StrTmp, "TBLINDEXMASTER")
End Sub

Private Sub CreateCATDESC()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[catdesc] (" & _
                " [cat] [nvarchar] (3) PRIMARY KEY, [Desc] [nvarchar] (50) NULL ," & _
                " [Dys] [nvarchar] (6) NULL ,[lt_allow] [real] NULL ," & _
                " [lt_Ignore] [real] NULL   ,[Erl_Allow] [real] NULL ," & _
                " [Erl_Ignore] [real] NULL  ,[Hf_cof] [real] NULL ," & _
                " [Fl_Cof] [real] NULL ,     [hf_cof_o] [real] NULL ," & _
                " [fl_cof_o] [real] NULL ,   [halfcutlt] [real] NULL ," & _
                " [halfcuter] [real] NULL ,  [ltinmnth] [real] NULL ," & _
                " [erinmnth] [real] NULL ,   [letcut] [real] NULL ," & _
                " [erlcut] [real] NULL ,     [everlet] [real] NULL ," & _
                " [evererl] [real] NULL ,    [fstletpr] [nvarchar] (2) NULL ," & _
                " [secletpr] [nvarchar] (2) NULL ,[trdletpr] [nvarchar] (2) NULL ," & _
                " [fsterlpr] [nvarchar] (2) NULL ,[secerlpr] [nvarchar] (2) NULL ," & _
                " [trderlpr] [nvarchar] (2) NULL ,[dederl] [nvarchar] (2) NULL ," & _
                " [dedlet] [nvarchar] (2) NULL , [laterule] [nvarchar] (2) NULL ," & _
                " [earlrule] [nvarchar] (2) NULL, [invisible] [nvarchar](2))"
    Case 3 ''Oracle
        StrTmp = "Create Table Catdesc ( cat Varchar2(3) primary key, ""Desc"" Varchar2(50)," & _
                "Dys Varchar2(6),lt_allow number(4,2),lt_Ignore number(4,2)," & _
                " Erl_Allow number(4,2), Erl_Ignore number(4,2), Hf_cof number(4,2)," & _
                " Fl_Cof number(4,2), hf_cof_o number(4,2), fl_cof_o number(4,2)," & _
                " halfcutlt number(4,2), halfcuter number(4,2), ltinmnth number(4,2)," & _
                " erinmnth number(4,2), letcut number(4,2), erlcut number(4,2)," & _
                " everlet number(4,2), evererl number(4,2), fstletpr Varchar2(2)," & _
                " secletpr Varchar2(2), trdletpr Varchar2(2), fsterlpr Varchar2(2)," & _
                " secerlpr Varchar2(2), trderlpr Varchar2(2), dederl Varchar2(2)," & _
                " dedlet Varchar2(2), laterule Varchar2(2), earlrule Varchar2(2), invisible varchar2(2))"
End Select
Call ExecuteConn(StrTmp, "CATDESC")
End Sub

Private Sub CreateCOMPANY()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[company] ( [Company] [tinyint] PRIMARY KEY ," & _
                " [CName] [nvarchar] (200) NULL)"
    Case 3 ''Oracle
        StrTmp = "Create Table Company(Company Number(2)primary key,CName Varchar2(200))"
End Select
Call ExecuteConn(StrTmp, "Company")
End Sub
Private Sub CreateCORUL()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[CORul] ([COCode] [tinyint] PRIMARY KEY ," & _
                " [CODesc] [nvarchar] (50) NULL ,[COWD] [tinyint] NULL ," & _
                " [COWO] [tinyint] NULL ,[COHL] [tinyint] NULL ," & _
                " [COAvail] [tinyint] NULL ,[WDH] [real] NULL ," & _
                " [WOH] [real] NULL ,[HLH] [real] NULL ,[WDF] [real] NULL ," & _
                " [WOF] [real] NULL ,[HLF] [real] NULL ,[DedLate] [tinyint] NULL ," & _
                " [DedEarl] [tinyint] NULL)"
    Case 3 ''Oracle
        StrTmp = "Create Table CORul(COCode Number(3)primary key,CODesc Varchar2(50),COWD Number(2)," & _
                " COWO Number(2),COHL Number(2),COAvail Number(2),WDH Number(4,2)," & _
                " WOH  Number(4,2),HLH  Number(4,2),WDF  Number(4,2),WOF  Number(4,2)," & _
                " HLF  Number(4,2),DedLate Number(2),DedEarl Number(2))"
End Select
Call ExecuteConn(StrTmp, "CORul")
End Sub
Private Sub CreateDAILYPRO()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[DailyPro] ([card] [nvarchar] (8) NULL ," & _
                " [empcode] [nvarchar] (8) NULL ,[Dte] [smalldatetime] NULL ," & _
                " [t_punch] [real] NULL ,[shift] [nvarchar] (3) NULL ," & _
                " [entry] [real] NULL ,[flg] [nvarchar] (2) NULL ," & _
                " [UnqFld] [int] IDENTITY (1,1) NOT NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table DailyPro(card Varchar2(8),empcode Varchar2(8)," & _
                "Dte Date,t_punch Number(4,2),shift Varchar2(3),entry Number(4,2)," & _
                "flg Varchar2(2),UnqFld Number(5))"
End Select
Call ExecuteConn(StrTmp, "DailyPro")
End Sub
Private Sub CreateDECLWOHL()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[declwohl] (" & _
                " [cat] [nvarchar] (3) NULL ,[date] [smalldatetime] NULL ," & _
                " [desc] [nvarchar] (50) NULL ,[hcode] [nvarchar] (8) NULL ," & _
                " [otrate] [real] NULL ,[compensdt] [smalldatetime] NULL ," & _
                " [declas] [nvarchar] (3) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table Declwohl(cat Varchar2(3),""Date"" Date,""Desc"" Varchar2(50)," & _
        "hcode Varchar2(8),otrate Number(4,2),compensdt Date,declas Varchar2(3))"
End Select
Call ExecuteConn(StrTmp, "Declwohl")
End Sub

Private Sub CreateDEPTDESC()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[deptdesc] ( [dept] [smallint] PRIMARY KEY ," & _
                 " [desc] [nvarchar] (30) NULL ,[strenth] [smallint] NULL," & _
                 " ) "
    Case 3 ''Oracle
        StrTmp = "Create Table DeptDesc(Dept Number(5) PRIMARY KEY,""Desc"" Varchar(50)," & _
                " strenth Number(3),EmailID Varchar2(15))"
End Select
Call ExecuteConn(StrTmp, "DeptDesc")
End Sub
Private Sub CreateDIVISION()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[Division] ( [div] [smallint]PRIMARY KEY ," & _
                 " [Divdesc] [nvarchar] (50) NULL ) "
    Case 3 ''Oracle
        StrTmp = "Create Table Division(Div Number(5) primary key,DivDesc Varchar2(50))"
End Select
Call ExecuteConn(StrTmp, "Division")
End Sub
Private Sub CreateDPERF()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[DPerf] ([Empcode] [nvarchar] (8) NULL ," & _
                " [PresAbsStr] [nvarchar] (190) NULL ,[punches] [nvarchar] (120) NULL" & _
                " ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table DPerf(Empcode Varchar2(8),PresAbsStr Varchar2(190)," & _
                " punches Varchar2(120))"
End Select
Call ExecuteConn(StrTmp, "DPerf")
End Sub
Private Sub CreateDPRAB()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[DprAb] ([Srno] [smallint] NULL ," & _
                " [Empcode] [nvarchar] (8) NULL ,[Present] [nvarchar] (6) NULL ," & _
                " [Absent] [nvarchar] (6) NULL ,[Offs] [nvarchar] (6) NULL ," & _
                " [OT] [float] NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table DPrAb(Srno Number(5),Empcode Varchar2(8)," & _
                " Present Varchar2(6),Absent Varchar2(6),Offs Varchar2(6),OT Number(5))"
End Select
Call ExecuteConn(StrTmp, "DPrAb")
End Sub
Private Sub CreateDSUM()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[DSum] ([Serial] [tinyint] NULL ," & _
                " [BaseC] [nvarchar] (50) NULL ,[Strength] [int] NULL ," & _
                " [NofE] [int] NULL ,[PP] [real] NULL ,[PPP] [real] NULL ," & _
                " [AA] [real] NULL ,[AAP] [real] NULL ,[WO] [real] NULL ," & _
                " [WOP] [real] NULL ,[HL] [real] NULL ,[HLP] [real] NULL ," & _
                " [PL] [real] NULL ,[PLP] [real] NULL ,[UPL] [real] NULL ," & _
                " [UPLP] [real] NULL ,[OD] [real] NULL ,[ODP] [real] NULL ," & _
                " [TOT] [real] NULL ,[TOTP] [real] NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table DSum(Serial Number(3),BaseC Varchar2(50)," & _
                " Strength Number(4),NofE Number(4),PP Number(6,2),PPP Number(6,2)," & _
                " AA Number(6,2),AAP Number(6,2),WO Number(6,2),WOP Number(6,2)," & _
                " HL Number(6,2),HLP Number(6,2),PL Number(6,2),PLP Number(6,2)," & _
                " UPL Number(6,2),UPLP Number(6,2),OD Number(6,2),ODP Number(6,2)," & _
                " TOT Number(6,2),TOTP Number(6, 2))"
End Select
Call ExecuteConn(StrTmp, "DSum")
End Sub
Private Sub CreateDTRN()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[Dtrn] ([empcode] [nvarchar] (8) NULL ," & _
                " [mndate] [smalldatetime] NULL ,[entry] [real] NULL ," & _
                " [entreq] [real] NULL ,[shift] [nvarchar] (3) NULL ," & _
                " [arrtim] [real] NULL ,[latehrs] [real] NULL ,[actrt_o] [real] NULL ," & _
                " [actrt_i] [real] NULL ,[actbreak] [real] NULL ,[deptim] [real] NULL ," & _
                " [earlhrs] [real] NULL ,[wrkhrs] [real] NULL ,[ovtim] [real] NULL ," & _
                " [time5] [real] NULL ,[time6] [real] NULL ,[time7] [real] NULL ," & _
                " [time8] [real] NULL ,[od_from] [real] NULL ,[od_to] [real] NULL ," & _
                " [ofd_from] [real] NULL ,[ofd_to] [real] NULL ,[ofd_hrs] [real] NULL ," & _
                " [present] [real] NULL ,[presabs] [nvarchar] (6) NULL ,[cof] [real] NULL ," & _
                " [chq] [nvarchar] (2) NULL ,[aflg] [nvarchar] (2) NULL ," & _
                " [dflg] [nvarchar] (2) NULL ,[ot_auth] [real] NULL ," & _
                " [et_hrs] [real] NULL ,[leave_hrs] [real] NULL ," & _
                " [ot_shift] [nvarchar] (6) NULL ,[Remarks] [nvarchar] (10) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table DTrn(empcode Varchar2(8),mndate Date,entry Number(2)," & _
                " entreq Number(2),shift Varchar2(3),arrtim Number(4,2)," & _
                " latehrs Number(4,2),actrt_o Number(4,2),actrt_i Number(4,2)," & _
                " actbreak Number(4,2),deptim Number(4,2),earlhrs Number(4,2)," & _
                " wrkhrs Number(4,2),ovtim Number(4,2),time5 Number(4,2)," & _
                " time6 Number(4,2),time7 Number(4,2),time8 Number(4,2)," & _
                " od_from Number(4,2),od_to Number(4,2),ofd_from Number(4,2)," & _
                " ofd_to Number(4,2),ofd_hrs Number(4,2),present Number(4,2)," & _
                " presabs Varchar2(6),cof Number(4,2),chq Varchar2(2)," & _
                " aflg Varchar2(2),dflg Varchar2(2),ot_auth Number(4,2)," & _
                " et_hrs Number(4,2),leave_hrs Number(4,2),ot_shift Varchar2(6)," & _
                " Remarks Varchar2(10))"
End Select
Call ExecuteConn(StrTmp, "DTrn")
End Sub
Private Sub CreateEMPFOUNDTB()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[empfoundtb] ([empcode] [nvarchar] (8) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table EmpFoundTB(Empcode Varchar2(8))"
End Select
Call ExecuteConn(StrTmp, "EmpFoundTB")
End Sub
Private Sub CreateEMPMST()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[Empmst] ([name] [nvarchar] (50) NULL ," & _
                " [designatn] [nvarchar] (50) NULL ,[empcode] [nvarchar] (8) PRIMARY KEY ," & _
                " [card] [nvarchar] (8) NULL ,[res_card] [nvarchar] (8) NULL ," & _
                " [resadd1] [nvarchar] (50) NULL ,[resadd2] [nvarchar] (50) NULL ," & _
                " [city] [nvarchar] (15) NULL ,[pin] [nvarchar] (15) NULL ," & _
                " [reference] [nvarchar] (50) NULL ,[phone] [nvarchar] (15) NULL ," & _
                " [joindate] [smalldatetime] NULL ,[confmdt] [smalldatetime] NULL ," & _
                " [salary] [real] NULL ,[earnsal] [real] NULL ,[addition] [real] NULL ," & _
                " [deduction] [real] NULL ,[paiddays] [real] NULL ,[div] [smallint] FOREIGN KEY REFERENCES division(div) ," & _
                " [dept] [smallint] FOREIGN KEY REFERENCES deptdesc(dept),[group] [smallint]FOREIGN KEY REFERENCES GroupMst([group]), [cat] [nvarchar] (3) FOREIGN KEY REFERENCES catdesc(cat) ," & _
                " [entry] [real] NULL ,[otflag] [bit] NULL ,[coflag] [bit] NULL ," & _
                " [styp] [nvarchar] (1) NULL ,[i_shif] [nvarchar] (6) NULL ," & _
                " [f_shf] [nvarchar] (3)FOREIGN KEY REFERENCES instshft(shift), [shift] [nvarchar] (3) NULL ," & _
                " [scode] [nvarchar] (3) FOREIGN KEY REFERENCES ro_shift(scode) ,[off] [nvarchar] (3) NULL ," & _
                " [off2] [nvarchar] (3) NULL ,[wo_1_3] [nvarchar] (3) NULL ," & _
                " [wo_2_4] [nvarchar] (3) NULL ,[weekoff] [nvarchar] (3) NULL ," & _
                " [p_time] [real] NULL ,[birth_dt] [smalldatetime] NULL ," & _
                " [bg] [nvarchar] (15) NULL ,[conv] [nvarchar] (1) NULL ," & _
                " [sex] [nvarchar] (1) NULL ,[qualf] [nvarchar] (10) NULL ," & _
                " [st] [nvarchar] (10) NULL ,[company] [tinyint] FOREIGN KEY REFERENCES company(company) ," & _
                " [abs_date] [smalldatetime] NULL ,[leavdate] [smalldatetime] NULL ," & _
                " [shf_date] [smalldatetime] NULL ,[shf_chg] [bit] NULL ," & _
                " [wrk_typ] [nvarchar] (10) NULL ,[wo] [nvarchar] (3) NULL ," & _
                " [email_id] [nvarchar] (40) NULL ,[baccount] [nvarchar] (15) NULL ," & _
                " [pfno] [nvarchar] (15) NULL ,[esino] [nvarchar] (15) NULL ,"
        StrTmp = StrTmp & "[esiflag] [nvarchar] (15) NULL ,[pfflag] [nvarchar] (15) NULL ," & _
                " [ca] [real] NULL ,[cca] [real] NULL ,[fda] [real] NULL ,[vda] [real] NULL ," & _
                " [hra] [real] NULL , [ea] [real] NULL ,[la] [real] NULL ,[medical] [real] NULL ," & _
                " [udf1] [nvarchar] (50) NULL ,[udf2] [nvarchar] (50) NULL ," & _
                " [udf3] [nvarchar] (50) NULL ,[udf4] [nvarchar] (50) NULL ," & _
                " [udf5] [nvarchar] (50) NULL ,[udf6] [nvarchar] (50) NULL ," & _
                " [udf7] [nvarchar] (50) NULL ,[udf8] [nvarchar] (50) NULL ," & _
                " [udf9] [nvarchar] (50) NULL ,[udf10] [nvarchar] (50) NULL ," & _
                " [OTCode] [tinyint] FOREIGN KEY REFERENCES OTRul(OTCode) ,[COCOde] [tinyint] FOREIGN KEY REFERENCES CORul(COcode) ," & _
                " [Location] [smallint] FOREIGN KEY REFERENCES Location(location) ,[Name2] [nvarchar] (50) NULL, " & _
                " [WOHLAction] [smallint] NULL,[Action3Shift] [nvarchar] (3) NULL," & _
                " [ActionBlank] [nvarchar] (3) NULL,[AutoForPunch] [smallint] NULL)"
    Case 3 ''Oracle
        StrTmp = "Create Table Empmst(name Varchar2(50),designatn Varchar2(50)," & _
                " empcode Varchar2(8)  primary key,card Varchar2(8),res_card Varchar2(8)," & _
                " resadd1 Varchar2(50),resadd2 Varchar2(50),city Varchar2(15)," & _
                " pin Varchar2(15),reference Varchar2(50),phone  Varchar2(15)," & _
                " joindate Date,confmdt Date,salary Number(7,2),earnsal Number(7,2)," & _
                " addition Number(7,2),deduction Number(7,2),paiddays Number(5,2)," & _
                " div Number(4)constraint myfkey11 references division(div),dept Number(5) constraint myfkey12 references deptdesc(dept),""Group"" Number(5) constraints myfkey13 references Groupmst(""Group""),cat Varchar2(30) constraint myfkey14 references catdesc(cat)," & _
                " entry   Number(2),otflag Number(3),coflag Number(3),styp Varchar2(1)," & _
                " i_shif Varchar2(6),f_shf Varchar2(3)constraint myfkey25 references instshft(shift),shift Varchar2(3),scode Varchar2(3) constraint myfkey15 references Ro_shift(scode)," & _
                " OFF Varchar2(3),off2 Varchar2(3),wo_1_3 Varchar2(3),wo_2_4 Varchar2(3)," & _
                " weekoff Varchar2(3),p_time Number(4,2),birth_dt Date,bg  Varchar2(15)," & _
                " conv Varchar2(1),sex Varchar2(1),qualf Varchar2(10),st Varchar2(10)," & _
                " company Number(2)constraint myfkey16 references company(company),abs_date Date,leavdate Date,shf_date Date," & _
                " shf_chg Number(1),wrk_typ Varchar2(10),wo Varchar2(3)," & _
                " email_id  Varchar2(40),baccount Varchar2(15),pfno Varchar2(15)," & _
                " esino Varchar2(15),esiflag Varchar2(15),pfflag Varchar2(15)," & _
                " ca Number(7,2),cca Number(7,2),fda Number(7,2),vda Number(7,2)," & _
                " hra Number(7,2),ea Number(7,2),la Number(7,2),medical Number(7,2)," & _
                " udf1 Varchar2(50),udf2 Varchar2(50),udf3 Varchar2(50)," & _
                " udf4 Varchar2(50),udf5 Varchar2(50),udf6 Varchar2(50)," & _
                " udf7 Varchar2(50),udf8 Varchar2(50),udf9 Varchar2(50)," & _
                " udf10 Varchar2(50),OTCode Number(3)constraint myfkey17 references OTRul(OTCode),COCOde Number(3)constraint myfkey18 references CORul(COCOde)," & _
                " Location Number(3)constraint myfkey19 references location(location),Name2 Varchar2(50),WOHLAction Number(3)," & _
                " Action3Shift Varchar2(3),ActionBlank Varchar2(3)," & _
                " AutoForPunch Number(3))"
End Select
Call ExecuteConn(StrTmp, "Empmst")
End Sub

Private Sub CreateEXC()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
       StrTmp = "CREATE TABLE [dbo].[Exc] ([Exclu] [int] NULL , [Daily] [int] NULL ,[Monthly] [int] NULL ,[Tmp] [nvarchar] (20) NULL ," & _
                " [Yearly] [int] NULL ,[UserNumber] [int] NULL ,[Lang] [nvarchar] (10) NULL" & _
                " ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table Exc(Exclu Number(2),Daily Number(2),Monthly Number(2)," & _
                " Tmp Varchar2(20),Yearly Number(2),UserNumber Number(4),Lang Varchar2(20))"
End Select
Call ExecuteConn(StrTmp, "Exc")
End Sub

Private Sub CreateGROUPMST()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[GroupMst] ([Group] [smallint] PRIMARY KEY ," & _
                " [GrupDesc] [nvarchar] (20) NULL)"
    Case 3 ''Oracle
        StrTmp = "Create Table GroupMst(""Group"" Number(5)primary key,GrupDesc Varchar2(50))"
End Select
Call ExecuteConn(StrTmp, "GroupMst")
End Sub

Private Sub CreateHOLIDAY()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[holiday] ([cat] [nvarchar] (3) FOREIGN KEY REFERENCES catdesc(cat)," & _
                " [Date] [smalldatetime] NULL ,[desc] [nvarchar] (50) NULL ," & _
                " [hcode] [nvarchar] (8) NULL ,[otrate] [real] NULL) "
    Case 3 ''Oracle
        StrTmp = "Create Table Holiday(cat Varchar2(3)constraint myfkey21 references catdesc(cat),""Date"" Date,""Desc"" Varchar2(50)," & _
                " hcode Varchar2(8),otrate Number(5, 2))"
End Select
Call ExecuteConn(StrTmp, "Holiday")
End Sub

Private Sub CreateINSTALL()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[install] ([dt_install] [smalldatetime] NULL ," & _
                " [american_dt] [bit] NULL  ,[upto] [real] NULL ," & _
                " [e_codesize] [real] NULL ,[e_cardsize] [real] NULL ," & _
                " [prsm_cards] [nvarchar] (1) NULL ,[cur_year] [nvarchar] (4) NULL ," & _
                " [hl_ot] [real] NULL ,[wo_ot] [real] NULL ,[ot_ot] [real] NULL ," & _
                " [prn_line] [real] NULL ,[PostLt] [real] NULL ,[PostErl] [real] NULL ," & _
                " [filt_time] [real] NULL ,[bus_card] [nvarchar] (8) NULL ," & _
                " [late_card] [nvarchar] (8) NULL ,[earl_card] [nvarchar] (8) NULL ," & _
                " [od_card] [nvarchar] (8) NULL ,[Tmp_card] [nvarchar] (8) NULL ," & _
                " [yearfrom] [real] NULL ,[weekfrom] [real] NULL,[pstart] [nvarchar] (8) NULL ," & _
                " [pend] [nvarchar] (8) NULL ,[lvupdtyear] [smallint] NULL ," & _
                " [allowedit] [real] NULL ,[deductlter] [bit] NULL  ," & _
                " [otround] [bit] NULL  ,[email] [bit] NULL  ,[definod] [bit] NULL  ," & _
                " [datpath] [nvarchar] (75) NULL ,[dec1] [real] NULL ,[dec1a] [real] NULL ," & _
                " [dec2] [real] NULL ,[dec2a] [real] NULL ,[dec3] [real] NULL ," & _
                " [dec3a] [real] NULL ,[dec4] [real] NULL ,[dec4a] [real] NULL ," & _
                " [dec5] [real] NULL ,[round1] [real] NULL ,[round2] [real] NULL ," & _
                " [round3] [real] NULL ,[round4] [real] NULL ,[round5] [real] NULL ," & _
                " [walpaper] [nvarchar] (75) NULL ,[defincut] [nvarchar] (1) NULL ," & _
                " [cutdt] [tinyint] NULL ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table Install(dt_install Date,american_dt Number(1)," & _
                " upto Number(4,2),e_codesize Number(1),e_cardsize Number(1)," & _
                " prsm_cards Varchar2(1),cur_year Varchar2(4),hl_ot Number(4,2)," & _
                " wo_ot Number(4,2),ot_ot Number(4,2),prn_line Number(4)," & _
                " PostLt Number(4,2),PostErl Number(4,2),filt_time Number(4,2)," & _
                " bus_card Varchar2(8),late_card Varchar2(8),earl_card Varchar2(8)," & _
                " od_card Varchar2(8),Tmp_card Varchar2(8),yearfrom Number(1)," & _
                " weekfrom Number(1),pstart Varchar2(8),pend Varchar2(8)," & _
                " lvupdtyear Number(4),allowedit  Number(4),deductlter Number(1)," & _
                " otround Number(1),email Number(1),definod Number(1)," & _
                " datpath Varchar2(75),dec1 Number(4,2),dec1a Number(4,2)," & _
                " dec2 Number(4,2),dec2a Number(4,2),dec3 Number(4,2),dec3a Number(4,2)," & _
                " dec4 Number(4,2),dec4a Number(4,2),dec5 Number(4,2),round1 Number(4,2)," & _
                " round2 Number(4,2),round3 Number(4,2),round4 Number(4,2)," & _
                " round5 Number(4,2),walpaper Varchar2(75),defincut Varchar2(1)," & _
                " cutdt Number(4))"

End Select
Call ExecuteConn(StrTmp, "Install")
End Sub

Private Sub CreateINSTSHFT()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[instshft] ([shift] [nvarchar] (3) PRIMARY KEY ," & _
                " [shf_in] [real] NULL ,[shf_out] [real] NULL ,[shf_hrs] [real] NULL ," & _
                " [rst_out] [real] NULL ,[rst_in] [real] NULL ,[rst_brk] [real] NULL ," & _
                " [rst_in_2] [real] NULL ,[rst_out_2] [real] NULL ,[rst_brk_2] [real] NULL ," & _
                " [rst_in_3] [real] NULL ,[rst_out_3] [real] NULL ,[rst_brk_3] [real] NULL ," & _
                " [night] [bit] NULL  ,[hdend] [real] NULL ,[hdstart] [real] NULL ," & _
                " [shiftname] [nvarchar] (50) NULL ,[brkshf] [nvarchar] (2) NULL )" & _
                " ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table InstShft(shift Varchar2(3) PRIMARY KEY,shf_in Number(4,2)," & _
                " shf_out Number(4,2),shf_hrs Number(4,2),rst_out Number(4,2)," & _
                " rst_in Number(4,2),rst_brk Number(4,2),rst_in_2 Number(4,2)," & _
                " rst_out_2 Number(4,2),rst_brk_2 Number(4,2),rst_in_3 Number(4,2)," & _
                " rst_out_3 Number(4,2),rst_brk_3 Number(4,2),night Number(1)," & _
                " hdend Number(4,2),hdstart Number(4,2),shiftname Varchar2(50)," & _
                " brkshf Varchar2(2))"
End Select
Call ExecuteConn(StrTmp, "InstShft")
End Sub

Private Sub CreateLATEERL()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[LateErl] ([EmpCode] [nvarchar] (8) NULL ," & _
                " [date] [smalldatetime] NULL ,[LateHrs] [real] NULL ," & _
                " [EarlHrs] [real] NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table LateErl(Empcode Varchar2(8),""Date"" Date," & _
                " LateHrs Number(4,2),EarlHrs Number(4, 2))"
End Select
Call ExecuteConn(StrTmp, "LateErl")
End Sub

Private Sub CreateLEAVBAL()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[leavbal] ([empcode] [nvarchar] (8) FOREIGN KEY REFERENCES empmst(empcode) " & _
                " )"
    Case 3 ''Oracle
        StrTmp = "Create Table Leavbal(Empcode Varchar2(8)constraint myfkey31 references empmst(empcode))"
End Select
Call ExecuteConn(StrTmp, "Leavbal")
End Sub

Private Sub CreateLEAVDESC()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[Leavdesc] ([leave] [nvarchar] (30) NULL ," & _
                " [lvcode] [nvarchar] (3) NULL ,[type] [nvarchar] (1) NULL ," & _
                " [paid] [nvarchar] (1) NULL ,[encase] [nvarchar] (1) NULL ," & _
                " [lv_cof] [nvarchar] (1) NULL ,[lv_qty] [real] NULL ," & _
                " [lv_acumul] [real] NULL ,[run_wrk] [nvarchar] (1) NULL ," & _
                " [isitleave] [nvarchar] (1) NULL ,[cat] [nvarchar] (3) FOREIGN KEY REFERENCES catdesc(cat) ," & _
                " [creditnow] [nvarchar] (1) NULL ,[fulcredit] [nvarchar] (1) NULL ," & _
                " [no_oftimes] [real] NULL ,[allowdays] [real] NULL ," & _
                " [minallowdays] [int] NULL ,[custcode] [nvarchar] (3) NULL )" & _
                " ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table LeavDesc(leave Varchar2(30),lvcode Varchar2(3)," & _
                " type Varchar2(1),paid Varchar2(1),encase Varchar2(1)," & _
                " lv_cof Varchar2(1),lv_qty Number(5,2),lv_acumul Number(5,2)," & _
                " run_wrk Varchar2(1),isitleave Varchar2(1),cat Varchar2(3)constraint myfkey41 references catdesc(cat)," & _
                " creditnow Varchar2(1),fulcredit Varchar2(1),no_oftimes Number(5,2)," & _
                " allowdays Number(4,2),minallowdays Number(4,2),custcode Varchar2(3))"
End Select
Call ExecuteConn(StrTmp, "LeavDesc")
End Sub

Private Sub CreateLEAVINFO()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[leavinfo] ([empcode] [nvarchar] (8) NULL ," & _
                " [trcd] [real] NULL ,[fromdate] [smalldatetime] NULL ," & _
                " [todate] [smalldatetime] NULL ,[lcode] [nvarchar] (3) NULL ," & _
                " [days] [float] NULL ,[lv_type_rw] [nvarchar] (1) NULL ," & _
                " [hf_option] [nvarchar] (6) NULL ,[fordate] [smalldatetime] NULL ," & _
                " [entrydate] [smalldatetime] NULL ,[advance] [nvarchar] (4) NULL" & _
                " ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table LeavInfo(Empcode Varchar2(8),trcd Number(1)," & _
                " fromdate Date,todate Date,lcode  Varchar2(3),days Number(5,2)," & _
                " lv_type_rw Varchar2(1),hf_option Varchar2(6),fordate Date," & _
                " entrydate Date,advance VarChar(4))"
End Select
Call ExecuteConn(StrTmp, "LeavInfo")
End Sub

Private Sub CreateLEAVTRN()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[leavtrn] ([empcode] [nvarchar] (8) NULL ," & _
                " [lst_date] [smalldatetime] NULL ,[paiddays] [real] NULL ," & _
                " [ot_hrs] [float] NULL ,[otpd_hrs] [float] NULL ,[actovtpd] [float] NULL ," & _
                " [lt_no] [real] NULL ,[lt_hrs] [float] NULL ,[erl_no] [real] NULL ," & _
                " [erl_hrs] [float] NULL ,[wrk_hrs] [float] NULL ,[night] [real] NULL ," & _
                " [OTWO] [real] NULL ,[OTHL] [real] NULL ,[OTNO] [real] NULL ," & _
                " [A] [float] NULL ,[HL] [float] NULL ,[P] [float] NULL ,[WO] [float] NULL ," & _
                " ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table LeavTrn(Empcode Varchar2(8),lst_date Date," & _
                " paiddays Number(5,2),ot_hrs Number(6,2),otpd_hrs Number(6,2)," & _
                " actovtpd Number(6,2),lt_no Number(5,2),lt_hrs Number(6,2)," & _
                " erl_no Number(5,2),erl_hrs Number(6,2),wrk_hrs Number(6,2)," & _
                " night Number(5,2),A Number(5,2),HL Number(5,2),P Number(5,2)," & _
                " WO Number(5, 2))"
End Select
Call ExecuteConn(StrTmp, "LeavTrn")
End Sub

Private Sub CreateLOCATION()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[Location] ([Location] [smallint] PRIMARY KEY ," & _
                " [LocDesc] [nvarchar] (50) NULL)"
    Case 3 ''Oracle
        StrTmp = "Create Table Location(Location Number(3)primary key,LocDesc Varchar2(50))"
End Select
Call ExecuteConn(StrTmp, "Location")
End Sub

Private Sub CreateLOST()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[lost] ([empcode] [nvarchar] (8) NULL ," & _
                " [dy] [nvarchar] (2) NULL ,[mn] [nvarchar] (2) NULL ," & _
                " [hrs] [nvarchar] (5) NULL ,[min] [nvarchar] (5) NULL ," & _
                " [date] [smalldatetime] NULL ,[t_punch] [real] NULL ," & _
                " [shift] [nvarchar] (3) NULL ,[entry] [real] NULL ," & _
                " [flg] [nvarchar] (2) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table Lost(Empcode Varchar2(8),dy Varchar2(2)," & _
                " mn Varchar2(2),hrs Varchar2(2),min Varchar2(2),""Date"" Date," & _
                " t_punch Number(4,2),shift Varchar2(3),entry Number(2),flg Varchar2(2))"
End Select
Call ExecuteConn(StrTmp, "Lost")
End Sub
Private Sub CreateLVT()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[lvt] ([empcode] [nvarchar] (8) NULL ," & _
                " [lst_date] [smalldatetime] NULL ,[paiddays] [real] NULL ," & _
                " [ot_hrs] [float] NULL ,[otpd_hrs] [float] NULL ," & _
                " [actovtpd] [float] NULL ,[lt_no] [real] NULL ,[lt_hrs] [float] NULL ," & _
                " [erl_no] [real] NULL ,[erl_hrs] [float] NULL ,[wrk_hrs] [float] NULL ," & _
                " [night] [real] NULL ,[OTWO] [real] NULL ,[OTHL] [real] NULL ," & _
                " [OTNO] [real] NULL ,[A] [float] NULL ,[HL] [float] NULL ," & _
                " [P] [float] NULL ,[WO] [float] NULL ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table LVT(Empcode Varchar2(8),lst_date Date," & _
                " paiddays Number(5,2),ot_hrs Number(6,2),otpd_hrs Number(6,2)," & _
                " actovtpd Number(6,2),lt_no Number(5,2),lt_hrs Number(6,2)," & _
                " erl_no Number(5,2),erl_hrs Number(6,2),wrk_hrs Number(6,2)," & _
                " night Number(5,2),A Number(5,2),HL Number(5,2),P Number(5,2)," & _
                " WO Number(5, 2))"
End Select
Call ExecuteConn(StrTmp, "LVT")
End Sub
Private Sub CreateLVTRNPERMT()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[lvtrnpermt] ([empcode] [nvarchar] (8) NULL ," & _
                " [lst_date] [smalldatetime] NULL ,[paiddays] [real] NULL ," & _
                " [ot_hrs] [float] NULL ,[otpd_hrs] [float] NULL ," & _
                " [actovtpd] [float] NULL ,[lt_no] [real] NULL ," & _
                " [lt_hrs] [float] NULL ,[erl_no] [real] NULL ," & _
                " [erl_hrs] [float] NULL ,[wrk_hrs] [float] NULL ,[night] [real] NULL " & _
                " ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table LvTrnPermt(Empcode Varchar2(8),lst_date Date," & _
                " paiddays Number(4,2),ot_hrs Number(6,2),otpd_hrs Number(6,2)," & _
                " actovtpd Number(6,2),lt_no Number(4,2),lt_hrs Number(6,2)," & _
                " erl_no Number(4,2),erl_hrs Number(6,2),wrk_hrs Number(6,2)," & _
                " night Number(4, 2))"
End Select
Call ExecuteConn(StrTmp, "LvTrnPermt")
End Sub
Private Sub CreateMALE()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[MALE] ([Empcode] [nvarchar] (8) NULL ," & _
                " [Absent] [nvarchar] (6) NULL ,[LateNo] [nvarchar] (6) NULL ," & _
                " [LateHrs] [nvarchar] (6) NULL ,[EarlyNo] [nvarchar] (6) NULL ," & _
                " [EarlyHrs] [nvarchar] (6) NULL ,[trcd] [nvarchar] (10) NULL ," & _
                " [fromdate] [nvarchar] (15) NULL ,[todate] [nvarchar] (15) NULL ," & _
                " [lcode] [nvarchar] (3) NULL ,[days] [nvarchar] (3) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table MALE(Empcode Varchar2(8),Absent Varchar2(6)," & _
                " LateNo Varchar2(6),LateHrs Varchar2(6),EarlyNo Varchar2(6)," & _
                " EarlyHrs Varchar2(6),trcd Varchar2(10),fromdate Varchar2(15)," & _
                " todate Varchar2(15),lcode Varchar2(3),days Varchar2(3))"
End Select
Call ExecuteConn(StrTmp, "MALE")
End Sub
Private Sub CreateMATT()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[MAtt] ([Empcode] [nvarchar] (8) NULL ," & _
                " [MndateStr] [nvarchar] (200) NULL ,[PresAbsStr] [nvarchar] (200) NULL ," & _
                " [LeaveStr] [nvarchar] (200) NULL ,[PDaysStr] [nvarchar] (200) NULL ," & _
                " [OtStr] [nvarchar] (200) NULL ,[WrkStr] [nvarchar] (200) NULL ," & _
                " [Nightstr] [nvarchar] (200) NULL ,[LvVal] [nvarchar] (200) NULL ," & _
                " [DaysDed] [nvarchar] (200) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table MATT(Empcode Varchar2(8),MndateStr Varchar2(200)," & _
                " PresAbsStr Varchar2(200),LeaveStr Varchar2(200)," & _
                " PDaysStr Varchar2(200),OtStr Varchar2(200),WrkStr Varchar2(200)," & _
                " Nightstr Varchar2(200),LvVal Varchar2(200),DaysDed Varchar2(200))"
End Select
Call ExecuteConn(StrTmp, "MATT")
End Sub
Private Sub CreateMEMOTABLE()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[MemoTable] ([MemoNum] [tinyint] NULL ," & _
                " [MemoText] [nvarchar] (125) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table Memotable(MemoNum Number(1),MemoText Varchar2(125))"
End Select
Call ExecuteConn(StrTmp, "Memotable")
End Sub
Private Sub CreateMONTRN()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[Montrn] ([empcode] [nvarchar] (8) NULL ," & _
                " [date] [smalldatetime] NULL ,[entry] [real] NULL ," & _
                " [entreq] [real] NULL ,[shift] [nvarchar] (3) NULL ," & _
                " [arrtim] [real] NULL ,[latehrs] [real] NULL ,[actrt_o] [real] NULL ," & _
                " [actrt_i] [real] NULL ,[actbreak] [real] NULL ,[deptim] [real] NULL ," & _
                " [earlhrs] [real] NULL ,[wrkhrs] [real] NULL ,[ovtim] [real] NULL ," & _
                " [time5] [real] NULL ,[time6] [real] NULL ,[time7] [real] NULL ," & _
                " [time8] [real] NULL ,[od_from] [real] NULL ,[od_to] [real] NULL ," & _
                " [ofd_from] [real] NULL ,[ofd_to] [real] NULL ,[ofd_hrs] [real] NULL ," & _
                " [present] [real] NULL ,[presabs] [nvarchar] (4) NULL ,[cof] [real] NULL ," & _
                " [chq] [nvarchar] (1) NULL ,[aflg] [nvarchar] (2) NULL ," & _
                " [dflg] [nvarchar] (2) NULL ,[ot_auth] [real] NULL ,[et_hrs] [real] NULL ," & _
                " [leave_hrs] [real] NULL ,[ot_shift] [nvarchar] (6) NULL ," & _
                " [Remarks] [nvarchar] (10) NULL ,[OTConf] [nvarchar] (50) NULL, " & _
                " [OTRem] [nvarchar] (15) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table Montrn(Empcode Varchar2(8),""Date"" Date," & _
                " entry Number(2),entreq Number(2),shift Varchar2(3),arrtim Number(4,2)," & _
                " latehrs Number(4,2),actrt_o Number(4,2),actrt_i Number(4,2)," & _
                " actbreak Number(4,2),deptim Number(4,2),earlhrs Number(4,2)," & _
                " wrkhrs Number(4,2),ovtim    Number(4,2),time5   Number(4,2)," & _
                " time6  Number(4,2),time7    Number(4,2),time8    Number(4,2)," & _
                " od_from Number(4,2),od_to Number(4,2),ofd_from Number(4,2)," & _
                " ofd_to Number(4,2),ofd_hrs Number(4,2),present Number(4,2)," & _
                " presabs Varchar2(4),cof Number(4,2),chq Varchar2(1),aflg Varchar2(2)," & _
                " dflg Varchar2(2),ot_auth Number(6,2),et_hrs Number(6,2)," & _
                " leave_hrs Number(6,2),ot_shift Varchar2(6),Remarks Varchar2(10)," & _
                " OTConf Varchar2(50),OTRem Varchar2(15))"
End Select
Call ExecuteConn(StrTmp, "Montrn")
End Sub

Private Sub CreateMPERF()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[MPerf] ([Empcode] [nvarchar] (8) NULL ," & _
                " [date] [nvarchar] (200) NULL ,[ArrStr] [nvarchar] (200) NULL ," & _
                " [DepStr] [nvarchar] (200) NULL ,[LateStr] [nvarchar] (200) NULL ," & _
                " [EarlStr] [nvarchar] (200) NULL ,[WorkStr] [nvarchar] (200) NULL ," & _
                " [OTStr] [nvarchar] (200) NULL ,[PresAbsStr] [nvarchar] (200) NULL ," & _
                " [ShfStr] [nvarchar] (200) NULL ,[sumlate] [float] NULL ," & _
                " [sumearly] [float] NULL ,[sumwork] [float] NULL ,[sumOT] [float] NULL " & _
                ") ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table MPerf(Empcode Varchar2(8),""Date"" Varchar2(200)," & _
                " ArrStr Varchar2(200),DepStr Varchar2(200),LateStr Varchar2(200)," & _
                " EarlStr Varchar2(200),WorkStr Varchar2(200),OTStr Varchar2(200)," & _
                " PresAbsStr Varchar2(200),ShfStr Varchar2(200),sumlate Number(5,2)," & _
                " sumearly Number(5,2),sumwork Number(5,2),sumOT Number(5, 2))"
End Select
Call ExecuteConn(StrTmp, "MPerf")
End Sub

Private Sub CreateNEWCAPTIONS()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[NewCaptions] ([ID] [nvarchar] (7) NULL ," & _
                " [CaptEng] [nvarchar] (500) NULL ,[CaptOther] [nvarchar] (500) NULL" & _
                " ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table NewCaptions(ID Varchar2(7),CaptEng Varchar2(500)," & _
                " CaptOther Varchar2(500))"
End Select
Call ExecuteConn(StrTmp, "NewCaptions")
End Sub

Private Sub CreateOTRUL()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[OTRul] ([OTCode] [tinyint] PRIMARY KEY ," & _
                " [OTDesc] [nvarchar] (50) NULL ,[OTWD] [tinyint] NULL ," & _
                " [OTWO] [tinyint] NULL ,[OTHL] [tinyint] NULL ,[WDRates] [tinyint] NULL ," & _
                " [WORates] [tinyint] NULL ,[HLRates] [tinyint] NULL ,[Authorized] [nvarchar] (1) NULL ," & _
                " [MaxOT] [real] NULL ,[DedLate] [tinyint] NULL ,[DedEarl] [tinyint] NULL ," & _
                " [From1] [real] NULL ,[To1] [real] NULL ,[Deduct1] [real] NULL ," & _
                " [All1] [tinyint] NULL ,[From2] [real] NULL ,[To2] [real] NULL ," & _
                " [Deduct2] [real] NULL ,[All2] [tinyint] NULL ,[From3] [real] NULL ," & _
                " [To3] [real] NULL ,[Deduct3] [real] NULL ,[All3] [tinyint] NULL ," & _
                " [MoreThan] [real] NULL ,[Deduct4] [real] NULL ,[All4] [tinyint] NULL ," & _
                " [WODeduct] [tinyint] NULL ,[HLDeduct] [tinyint] NULL ," & _
                " [RFrom1] [real] NULL ,[RTo1] [real] NULL ,[Round1] [real] NULL ," & _
                " [RFrom2] [real] NULL ,[RTo2] [real] NULL ,[Round2] [real] NULL ," & _
                " [RFrom3] [real] NULL ,[RTo3] [real] NULL ,[Round3] [real] NULL ," & _
                " [RFrom4] [real] NULL ,[RTo4] [real] NULL ,[Round4] [real] NULL ," & _
                " [RTo5] [real] NULL ,[Round5] [real] NULL )"
    Case 3 ''Oracle
        StrTmp = "Create Table OTRul(OTCode Number(3)primary key,OTDesc Varchar2(50)," & _
                " OTWD Number(3),OTWO Number(3),OTHL Number(3),WDRates Number(5,2)," & _
                " WORates Number(5,2),HLRates Number(5,2),Authorized Varchar2(1)," & _
                " MaxOT Number(5,2),DedLate Number(2),DedEarl Number(2),From1 Number(4,2)," & _
                " To1 Number(4,2),Deduct1 Number(4,2),All1 Number(2),From2 Number(4,2)," & _
                " To2 Number(4,2),Deduct2 Number(4,2),All2 Number(2),From3 Number(4,2)," & _
                " To3 Number(4,2),Deduct3 Number(4,2),All3 Number(2),MoreThan Number(4,2)," & _
                " Deduct4 Number(4,2),All4 Number(2),WODeduct Number(2),HLDeduct Number(2)," & _
                " RFrom1 Number(4,2),RTo1 Number(4,2),Round1 Number(4,2),RFrom2 Number(4,2)," & _
                " RTo2 Number(4,2),Round2 Number(4,2),RFrom3 Number(4,2),RTo3 Number(4,2)," & _
                " Round3 Number(4,2),RFrom4 Number(4,2),RTo4 Number(4,2),Round4 Number(4,2)," & _
                " RTo5 Number(4,2),Round5 Number(4, 2))"
End Select
Call ExecuteConn(StrTmp, "OTRul")
End Sub

Private Sub CreatePPERF()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[PPerf] ([Empcode] [nvarchar] (8) NULL ," & _
                " [date] [nvarchar] (200) NULL ,[ArrStr] [nvarchar] (200) NULL ," & _
                " [DepStr] [nvarchar] (200) NULL ,[LateStr] [nvarchar] (200) NULL ," & _
                " [EarlStr] [nvarchar] (200) NULL ,[WorkStr] [nvarchar] (200) NULL ," & _
                " [OTStr] [nvarchar] (200) NULL ,[PresAbsStr] [nvarchar] (200) NULL ," & _
                " [ShfStr] [nvarchar] (200) NULL ,[sumLate] [float] NULL ," & _
                " [SumEarly] [float] NULL ,[SumWork] [float] NULL ," & _
                " [SumExtra] [float] NULL ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table PPerf(Empcode Varchar2(8),""Date"" Varchar2(200)," & _
                " ArrStr Varchar2(200),DepStr Varchar2(200),LateStr Varchar2(200)," & _
                " EarlStr Varchar2(200),WorkStr Varchar2(200),OTStr Varchar2(200)," & _
                " PresAbsStr Varchar2(200),ShfStr Varchar2(200),Actrt_o Varchar2(200)," & _
                " Actrt_i Varchar2(200),Time5 Varchar2(200),Time6 Varchar2(200)," & _
                " Time7 Varchar2(200),Time8 Varchar2(200),sumLate Number(5,2)," & _
                " SumEarly Number(5,2),SumWork Number(5,2),SumExtra Number(5,2))"
End Select
Call ExecuteConn(StrTmp, "PPerf")
End Sub

Private Sub CreateRO_SHIFT()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[ro_shift] ([scode] [nvarchar] (3) PRIMARY KEY," & _
                " [name] [nvarchar] (50) NULL ,[skp] [nvarchar] (93) NULL ," & _
                " [pattern] [nvarchar] (93) NULL ,[mon_oth] [nvarchar] (1) NULL ," & _
                " [tot_shf] [real] NULL ,[tot_skp] [real] NULL ,[day_skp] [real] NULL" & _
                " ) "
    Case 3 ''Oracle
        StrTmp = "Create Table Ro_Shift(scode Varchar2(3)primary key,name Varchar2(50)," & _
                " skp Varchar2(93),pattern Varchar2(93),mon_oth Varchar2(1)," & _
                " tot_shf Number(5,2),tot_skp Number(5,2),day_skp Number(5, 2))"
End Select
Call ExecuteConn(StrTmp, "Ro_Shift")
End Sub

Private Sub CreateSHFINFO()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[shfinfo] ([empcode] [nvarchar] (8) NULL ," & _
                " [d1] [nvarchar] (3) NULL ,[d2] [nvarchar] (3) NULL ," & _
                " [d3] [nvarchar] (3) NULL ,[d4] [nvarchar] (3) NULL ," & _
                " [d5] [nvarchar] (3) NULL ,[d6] [nvarchar] (3) NULL ," & _
                " [d7] [nvarchar] (3) NULL ,[d8] [nvarchar] (3) NULL ," & _
                " [d9] [nvarchar] (3) NULL ,[d10] [nvarchar] (3) NULL ," & _
                " [d11] [nvarchar] (3) NULL ,[d12] [nvarchar] (3) NULL ," & _
                " [d13] [nvarchar] (3) NULL ,[d14] [nvarchar] (3) NULL ," & _
                " [d15] [nvarchar] (3) NULL ,[d16] [nvarchar] (3) NULL ," & _
                " [d17] [nvarchar] (3) NULL ,[d18] [nvarchar] (3) NULL ," & _
                " [d19] [nvarchar] (3) NULL ,[d20] [nvarchar] (3) NULL ," & _
                " [d21] [nvarchar] (3) NULL ,[d22] [nvarchar] (3) NULL ," & _
                " [d23] [nvarchar] (3) NULL ,[d24] [nvarchar] (3) NULL ," & _
                " [d25] [nvarchar] (3) NULL ,[d26] [nvarchar] (3) NULL ," & _
                " [d27] [nvarchar] (3) NULL ,[d28] [nvarchar] (3) NULL ," & _
                " [d29] [nvarchar] (3) NULL ,[d30] [nvarchar] (3) NULL ," & _
                " [d31] [nvarchar] (3) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table shfinfo(Empcode Varchar2(8),d1 Varchar2(3),d2 Varchar2(3)," & _
                " d3 Varchar2(3),d4 Varchar2(3),d5 Varchar2(3),d6 Varchar2(3)," & _
                " d7 Varchar2(3),d8 Varchar2(3),d9 Varchar2(3),d10 Varchar2(3)," & _
                " d11 Varchar2(3),d12 Varchar2(3),d13 Varchar2(3),d14 Varchar2(3)," & _
                " d15 Varchar2(3),d16 Varchar2(3),d17 Varchar2(3),d18 Varchar2(3)," & _
                " d19 Varchar2(3),d20 Varchar2(3),d21 Varchar2(3),d22 Varchar2(3)," & _
                " d23 Varchar2(3),d24 Varchar2(3),d25 Varchar2(3),d26 Varchar2(3)," & _
                " d27 Varchar2(3),d28 Varchar2(3),d29 Varchar2(3),d30 Varchar2(3)," & _
                " d31 Varchar2(3))"
End Select
Call ExecuteConn(StrTmp, "shfinfo")
End Sub

Private Sub CreateTBLDATA()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[tblData] ([strF1] [nvarchar] (20) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table TblData(strF1 Varchar2(20))"
End Select
Call ExecuteConn(StrTmp, "TblData")
End Sub

Private Sub CreateUSERACCS()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[UserAccs] ([UserName] [nvarchar] (20) NULL," & _
        "[Password] [nvarchar] (20) NULL,[UserType] [nvarchar] (20) NULL," & _
        "[DEPT] [nvarchar] (500) NULL,[HODRights] [nvarchar] (60) NULL," & _
        "[MasterRights] [nvarchar] (60) NULL,[LeaveRights] [nvarchar] (16) NULL," & _
        "[OtherRights1] [nvarchar] (60) NULL,[OtherRights2] [nvarchar] (60) NULL," & _
        "[OtherPass1] [nvarchar] (20) NULL,[OtherPass2] [nvarchar] (20) NULL," & _
        "[UserAddDate] [smalldatetime] NULL,[UserAddUser] [nvarchar] (20) NULL," & _
        "[UserModDate] [smalldatetime] NULL,[UserModdUser] [nvarchar] (20) NULL " & _
        " ) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table UserAccs(UserName Varchar2(20),Password Varchar2(20)," & _
                " UserType Varchar2(20),Dept  Varchar2(500),HODRights Varchar2(60)," & _
                " MasterRights Varchar2(60),LeaveRights Varchar2(16)," & _
                " OtherRights1 Varchar2(60),OtherRights2 Varchar2(60)," & _
                " OtherPass1 Varchar2(20),OtherPass2 Varchar2(20),UserAddDate Date," & _
                " UserAddUser Varchar2(20),UserModDate Date,UserModUser Varchar2(20))"
End Select
Call ExecuteConn(StrTmp, "UserAccs")
End Sub

'Private Sub createWPerf()
'Dim StrTmp As String
'Select Case TDSN.BackEnd
'    Case 1 ''Sql-Server
'        StrTmp = "CREATE TABLE [dbo].[WPerf] ([Empcode] [nvarchar] (8) NULL ," & _
'                " [date] [nvarchar] (200) NULL ,[ArrStr] [nvarchar] (200) NULL ," & _
'                " [DepStr] [nvarchar] (200) NULL ,[LateStr] [nvarchar] (200) NULL ," & _
'                " [EarlStr] [nvarchar] (200) NULL ,[WorkStr] [nvarchar] (200) NULL ," & _
'                " [OTStr] [nvarchar] (200) NULL ,[PresAbsStr] [nvarchar] (200) NULL ," & _
'                " [ShfStr] [nvarchar] (200) NULL , [sumLate] [float] NULL ," & _
'                " [SumEarly] [float] NULL ,[SumWork] [float] NULL ," & _
'                " [SumExtra] [float] NULL ,[punches] [nvarchar] (255) NULL " & _
'                " ) ON [PRIMARY]"
'    Case 3 ''Oracle
'        StrTmp = "Create Table WPerf(Empcode Varchar2(8),""Date"" Varchar2(200)," & _
'                " ArrStr Varchar2(200),DepStr Varchar2(200),LateStr Varchar2(200)," & _
'                " EarlStr Varchar2(200),WorkStr Varchar2(200),OTStr Varchar2(200)," & _
'                " PresAbsStr Varchar2(200),ShfStr Varchar2(200),sumLate Number(5,2)," & _
'                " SumEarly Number(5,2),SumWork Number(5,2),SumExtra Number(5,2)," & _
'                " Punches Varchar2(255))"
'End Select
'Call ExecuteConn(StrTmp, "WPerf")
'End Sub
Private Sub CreateWSTAT()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[WStat] ([empCode] [nvarchar] (8) NULL ," & _
                " [FrW] [nvarchar] (6) NULL ,[SecW] [nvarchar] (6) NULL ," & _
                " [thW] [nvarchar] (6) NULL ,[FoW] [nvarchar] (6) NULL ," & _
                " [FiW] [nvarchar] (6) NULL ,[SiW] [nvarchar] (6) NULL ," & _
                " [SevW] [nvarchar] (6) NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table WStat(EmpCode Varchar2(8),FrW Varchar2(6)," & _
                " SecW Varchar2(6),thW Varchar2(6),FoW Varchar2(6),FiW Varchar2(6)," & _
                " SiW Varchar2(6),SevW Varchar2(6))"
End Select
Call ExecuteConn(StrTmp, "WStat")
End Sub
Private Sub CreateYRTB()
Dim StrTmp As String
Select Case TDSN.BackEnd
    Case 1 ''Sql-Server
        StrTmp = "CREATE TABLE [dbo].[YrTb] ([Empcode] [nvarchar] (8) NULL ," & _
                " [YStr] [nvarchar] (6) NULL ,[YValStr] [nvarchar] (200) NULL ," & _
                " [PdDaysStr] [nvarchar] (10) NULL ,[WrkStr] [nvarchar] (10) NULL ," & _
                " [NightStr] [nvarchar] (10) NULL ,[LtNo] [nvarchar] (10) NULL ," & _
                " [Latehrs] [nvarchar] (10) NULL ,[ErNo] [nvarchar] (10) NULL ," & _
                " [EarlHrs] [nvarchar] (10) NULL ,[FromLv] [nvarchar] (10) NULL ," & _
                " [CreditLv] [nvarchar] (10) NULL ,[AvailLv] [nvarchar] (10) NULL ," & _
                " [ToDate] [smalldatetime] NULL ,[Lcode] [nvarchar] (3) NULL ," & _
                " [fromdate] [smalldatetime] NULL ,[trcd] [nvarchar] (50) NULL ," & _
                " [counter] [int] NULL) ON [PRIMARY]"
    Case 3 ''Oracle
        StrTmp = "Create Table YrTb(Empcode Varchar2(8),YStr Varchar2(6)," & _
                " YValStr Varchar2(200),PdDaysStr Varchar2(10),WrkStr Varchar2(10)," & _
                " NightStr Varchar2(10),LtNo Varchar2(10),Latehrs Varchar2(10)," & _
                " ErNo Varchar2(10),EarlHrs Varchar2(10),FromLv Varchar2(10)," & _
                " CreditLv Varchar2(10),AvailLv Varchar2(10),ToDate Date," & _
                " Lcode Varchar2(3),fromdate Date,trcd Varchar2(50),Counter Number(5))"
End Select
Call ExecuteConn(StrTmp, "YrTb")
End Sub

Private Sub CreateSequence()
Dim StrTmp As String
StrTmp = "Create Sequence NEXT1 Cache 99998 Maxvalue 99999 Cycle"
Call ExecuteConn(StrTmp, "NEXT1")
End Sub

Private Sub ExecuteConn(ByVal strQ As String, strTName As String)
On Error GoTo Err_P
Dim StrTmp As String
frmCreate.lblDisp.Caption = "Creating => " & strTName
frmCreate.lblDisp.Refresh
VstarConn.Execute strQ
frmCreate.lblDisp.Caption = ""
frmCreate.lblDisp.Refresh
Exit Sub
Err_P:
    If Err.Number = -2147217900 Then
        ''Already Exists then continue
    Else
        StrTmp = "There is some problem creating Table/Sequence " & strTName & vbCrLf & _
        "Do you wish to continue?" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
        "Please check for the following ..." & vbCrLf & vbTab & _
        "-> Database connectivity is no longer available." & vbCrLf & vbTab & "-> The User '" & _
        TDSN.UserName & "' does not have enough previleges for Table/Sequence creation."
        If MsgBox(StrTmp, vbYesNo + vbQuestion) = vbNo Then
            End
        End If
    End If
End Sub


Public Function InsertOracleCaptions() As Boolean
On Error GoTo Err_P
VstarConn.Execute "Delete from Newcaptions"
VstarConn.Execute "Commit"
VstarConn.Execute " Insert Into NewCaptions Values('00001','Not Enough Rights, Access Denied','00001') "
VstarConn.Execute " Insert Into NewCaptions Values('00002','&' || 'OK','00002') "
VstarConn.Execute " Insert Into NewCaptions Values('00003','&' || 'Cancel','00003') "
VstarConn.Execute " Insert Into NewCaptions Values('00004','&' || 'Add','00004') "
VstarConn.Execute " Insert Into NewCaptions Values('00005','&' || 'Edit','00005') "
VstarConn.Execute " Insert Into NewCaptions Values('00006','&' || 'Delete','00006') "
VstarConn.Execute " Insert Into NewCaptions Values('00007','&' || 'Save','00007') "
VstarConn.Execute " Insert Into NewCaptions Values('00008','E&' || 'xit','00008') "
VstarConn.Execute " Insert Into NewCaptions Values('00009','Continue ?','00009') "
VstarConn.Execute " Insert Into NewCaptions Values('00010','From','00010') "
VstarConn.Execute " Insert Into NewCaptions Values('00011','To','00011') "
VstarConn.Execute " Insert Into NewCaptions Values('00012','Days','00012') "
VstarConn.Execute " Insert Into NewCaptions Values('00013','List','00013') "
VstarConn.Execute " Insert Into NewCaptions Values('00014','Details','00014') "
VstarConn.Execute " Insert Into NewCaptions Values('00015','Are you sure to Delete this Record ?','00015') "
VstarConn.Execute " Insert Into NewCaptions Values('00016','From Date cannot be Empty.','00016') "
VstarConn.Execute " Insert Into NewCaptions Values('00017','To Date cannot be Empty.','00017') "
VstarConn.Execute " Insert Into NewCaptions Values('00018','From Date cannot be greater than To Date.','00018') "
VstarConn.Execute " Insert Into NewCaptions Values('00019','From Date','00019') "
VstarConn.Execute " Insert Into NewCaptions Values('00020','To Date','00020') "
VstarConn.Execute " Insert Into NewCaptions Values('00021',' not in Current Year.','00021') "
VstarConn.Execute " Insert Into NewCaptions Values('00022','&' || 'Close','00022') "
VstarConn.Execute " Insert Into NewCaptions Values('00023','Hours','00023') "
VstarConn.Execute " Insert Into NewCaptions Values('00024','Minutes cannot be greater than 0.59','00024') "
VstarConn.Execute " Insert Into NewCaptions Values('00025','Maximum value cannot be greater than 23.59','00025') "
VstarConn.Execute " Insert Into NewCaptions Values('00026','Month','00026') "
VstarConn.Execute " Insert Into NewCaptions Values('00027','Year','00027') "
VstarConn.Execute " Insert Into NewCaptions Values('00028','&' || 'Month','00028') "
VstarConn.Execute " Insert Into NewCaptions Values('00029','&' || 'Year','00029') "
VstarConn.Execute " Insert Into NewCaptions Values('00030','Date','00030') "
VstarConn.Execute " Insert Into NewCaptions Values('00031','Shift','00031') "
VstarConn.Execute " Insert Into NewCaptions Values('00032','Entry','00032') "
VstarConn.Execute " Insert Into NewCaptions Values('00033','Status','00033') "
VstarConn.Execute " Insert Into NewCaptions Values('00034','Arrival','00034') "
VstarConn.Execute " Insert Into NewCaptions Values('00035','Late','00035') "
VstarConn.Execute " Insert Into NewCaptions Values('00036','Departure','00036') "
VstarConn.Execute " Insert Into NewCaptions Values('00037','Early','00037') "
VstarConn.Execute " Insert Into NewCaptions Values('00038','Overtime','00038') "
VstarConn.Execute " Insert Into NewCaptions Values('00039','&' || 'Finish','00039') "
VstarConn.Execute " Insert Into NewCaptions Values('00040','Select &' || 'Range','00040') "
VstarConn.Execute " Insert Into NewCaptions Values('00041','&' || 'Unselect Range','00041') "
VstarConn.Execute " Insert Into NewCaptions Values('00042','&' || 'Select All','00042') "
VstarConn.Execute " Insert Into NewCaptions Values('00043','U&' || 'nselect All','00043') "
VstarConn.Execute " Insert Into NewCaptions Values('00044','Employee Selection','00044') "
VstarConn.Execute " Insert Into NewCaptions Values('00045','Fro&' || 'm','00045') "
VstarConn.Execute " Insert Into NewCaptions Values('00046','T&' || 'o','00046') "
VstarConn.Execute " Insert Into NewCaptions Values('00047','Code','00047') "
VstarConn.Execute " Insert Into NewCaptions Values('00048','Name','00048') "
VstarConn.Execute " Insert Into NewCaptions Values('00049','Please Select the Employees','00049') "
VstarConn.Execute " Insert Into NewCaptions Values('00050','Please Select the Dat File.','00050') "
VstarConn.Execute " Insert Into NewCaptions Values('00051','Category','00051') "
VstarConn.Execute " Insert Into NewCaptions Values('00052','Description','00052') "
VstarConn.Execute " Insert Into NewCaptions Values('00053','&' || 'Create','00053') "
VstarConn.Execute " Insert Into NewCaptions Values('00054','Leave Transaction File for the year.','00054') "
VstarConn.Execute " Insert Into NewCaptions Values('00055',' not found.','00055') "
VstarConn.Execute " Insert Into NewCaptions Values('00056','Decimal value can be 0.5 or 0 Only','00056') "
VstarConn.Execute " Insert Into NewCaptions Values('00057','Company','00057') "
VstarConn.Execute " Insert Into NewCaptions Values('00058','Department','00058') "
VstarConn.Execute " Insert Into NewCaptions Values('00059','Group','00059') "
VstarConn.Execute " Insert Into NewCaptions Values('00060','Minimum Value cannot be less than 0.','00060') "
VstarConn.Execute " Insert Into NewCaptions Values('00061','Employee Code','00061') "
VstarConn.Execute " Insert Into NewCaptions Values('00062','Data is  in Process : Please Try after Some Time','00062') "
VstarConn.Execute " Insert Into NewCaptions Values('00063','Leave','00063') "
VstarConn.Execute " Insert Into NewCaptions Values('00064','&' || 'Reset','00064') "
VstarConn.Execute " Insert Into NewCaptions Values('00065','Mon','00065') "
VstarConn.Execute " Insert Into NewCaptions Values('00066','Tue','00066') "
VstarConn.Execute " Insert Into NewCaptions Values('00067','Wed','00067') "
VstarConn.Execute " Insert Into NewCaptions Values('00068','Thu','00068') "
VstarConn.Execute " Insert Into NewCaptions Values('00069','Fri','00069') "
VstarConn.Execute " Insert Into NewCaptions Values('00070','Sat','00070') "
VstarConn.Execute " Insert Into NewCaptions Values('00071','Sun','00071') "
VstarConn.Execute " Insert Into NewCaptions Values('00072','Date cannot be blank','00072') "
VstarConn.Execute " Insert Into NewCaptions Values('00073','Non-Leap Year cannot have 29 days in February','00073') "
VstarConn.Execute " Insert Into NewCaptions Values('00074','Selec&' || 't Printer','00074') "
VstarConn.Execute " Insert Into NewCaptions Values('00075','&' || 'Send','00075') "
VstarConn.Execute " Insert Into NewCaptions Values('00076','Pre&' || 'view','00076') "
VstarConn.Execute " Insert Into NewCaptions Values('00077','&' || 'Print','00077') "
VstarConn.Execute " Insert Into NewCaptions Values('00078','&' || 'File','00078') "
VstarConn.Execute " Insert Into NewCaptions Values('00079','No Records found.','00079') "
VstarConn.Execute " Insert Into NewCaptions Values('00080','You have been marked absent on following dates. Please fill up OD/LEAVES at the earliest','00080') "
VstarConn.Execute " Insert Into NewCaptions Values('00081','You have been marked late on the following dates. Kindly forward OD/REQUISITE PERMISSION within 3 days','00081') "
VstarConn.Execute " Insert Into NewCaptions Values('00082','You have been marked early on following dates. Kindly forward ON/LEAVES/REQUISITE PERMISSION at the earliest','00082') "
VstarConn.Execute " Insert Into NewCaptions Values('00083','Login','00083') "
VstarConn.Execute " Insert Into NewCaptions Values('00084','Password','00084') "
VstarConn.Execute " Insert Into NewCaptions Values('00085','Invalid User','00085') "
VstarConn.Execute " Insert Into NewCaptions Values('00086','Incorrect PassWord','00086') "
VstarConn.Execute " Insert Into NewCaptions Values('00087','One or More Required Yearly Leave Files Missing :: Please Re-Create','00087') "
VstarConn.Execute " Insert Into NewCaptions Values('00088','OT Rules','00088') "
VstarConn.Execute " Insert Into NewCaptions Values('00089','CO Rules','00089') "
VstarConn.Execute " Insert Into NewCaptions Values('00090','OT Rule','00090') "
VstarConn.Execute " Insert Into NewCaptions Values('00091','CO Rule','00091') "
VstarConn.Execute " Insert Into NewCaptions Values('00092','Weekdays','00092') "
VstarConn.Execute " Insert Into NewCaptions Values('00093','Weekoffs','00093') "
VstarConn.Execute " Insert Into NewCaptions Values('00094','Holidays','00094') "
VstarConn.Execute " Insert Into NewCaptions Values('00095','Present','00095') "
VstarConn.Execute " Insert Into NewCaptions Values('00096','Absent','00096') "
VstarConn.Execute " Insert Into NewCaptions Values('00097','Deductions','00097') "
VstarConn.Execute " Insert Into NewCaptions Values('00098','Deduct','00098') "
VstarConn.Execute " Insert Into NewCaptions Values('00099','All','00099') "
VstarConn.Execute " Insert Into NewCaptions Values('00100','YES','00100') "
VstarConn.Execute " Insert Into NewCaptions Values('00101','NO','00101') "
VstarConn.Execute " Insert Into NewCaptions Values('00102','Invalid User Name or Password','00102') "
VstarConn.Execute " Insert Into NewCaptions Values('00103','Password changed successfully','00103') "
VstarConn.Execute " Insert Into NewCaptions Values('00104','OT Authorization','00104') "
VstarConn.Execute " Insert Into NewCaptions Values('00105','Reports','00105') "
VstarConn.Execute " Insert Into NewCaptions Values('00106','Report','00106') "
VstarConn.Execute " Insert Into NewCaptions Values('00107','Unacceptable Password :: Please Type New Password','00107') "
VstarConn.Execute " Insert Into NewCaptions Values('00108','Please Enter Password','00108') "
VstarConn.Execute " Insert Into NewCaptions Values('00109','Locations','00109') "
VstarConn.Execute " Insert Into NewCaptions Values('00110','Location','00110') "
VstarConn.Execute " Insert Into NewCaptions Values('00111','Click to Toggle Selection','00111') "
VstarConn.Execute " Insert Into NewCaptions Values('00112','Operation Invalid before Employee Joindate','00112') "
VstarConn.Execute " Insert Into NewCaptions Values('00113','To timings cannot be less then from timings','00113') "
VstarConn.Execute " Insert Into NewCaptions Values('00114','&' || 'Abort','00114') "
VstarConn.Execute " Insert Into NewCaptions Values('00115','Are you Sure to Abort the Process ?','00115') "
VstarConn.Execute " Insert Into NewCaptions Values('00116','Unauthorized OT','00116') "
VstarConn.Execute " Insert Into NewCaptions Values('00117','Authorized','00117') "
VstarConn.Execute " Insert Into NewCaptions Values('00118','Unauthorized','00118') "
VstarConn.Execute " Insert Into NewCaptions Values('00119','Authorized OT','00119') "
VstarConn.Execute " Insert Into NewCaptions Values('00120','Father Name','00120') "
VstarConn.Execute " Insert Into NewCaptions Values('00121','Entry','00121') "
VstarConn.Execute " Insert Into NewCaptions Values('00122','Entries','00122') "
VstarConn.Execute " Insert Into NewCaptions Values('00123','Designation','00123') "
VstarConn.Execute " Insert Into NewCaptions Values('00124','Remarks','00124') "
VstarConn.Execute " Insert Into NewCaptions Values('00125','OT Remark','00125') "
VstarConn.Execute " Insert Into NewCaptions Values('00126','Division','00126') "
VstarConn.Execute " Insert Into NewCaptions Values('00127','Divisions','00127') "
VstarConn.Execute " Insert Into NewCaptions Values('00128','Referential Integrity Error. Cannot Delete this record.','00128') "
VstarConn.Execute " Insert Into NewCaptions Values('01001','Select Date','01001') "
VstarConn.Execute " Insert Into NewCaptions Values('01002','&' || 'OK','01002') "
VstarConn.Execute " Insert Into NewCaptions Values('01003','&' || 'Cancel','01003') "
VstarConn.Execute " Insert Into NewCaptions Values('02001','Change Password','02001') "
VstarConn.Execute " Insert Into NewCaptions Values('02002','Old Settings','02002') "
VstarConn.Execute " Insert Into NewCaptions Values('02003','Menu Item','02003') "
VstarConn.Execute " Insert Into NewCaptions Values('02004','Old Password','02004') "
VstarConn.Execute " Insert Into NewCaptions Values('02005','Settings','02005') "
VstarConn.Execute " Insert Into NewCaptions Values('02006','New Password','02006') "
VstarConn.Execute " Insert Into NewCaptions Values('02007','Reconfirm Password','02007') "
VstarConn.Execute " Insert Into NewCaptions Values('02008','Please Select the Option for which the Password is to be Changed','02008') "
VstarConn.Execute " Insert Into NewCaptions Values('02009','Please Enter Old Password','02009') "
VstarConn.Execute " Insert Into NewCaptions Values('02010','Please Enter New Password','02010') "
VstarConn.Execute " Insert Into NewCaptions Values('02011','Please Confirm New Password','02011') "
VstarConn.Execute " Insert Into NewCaptions Values('02012','Cannot confirm new Password. Please try again.','02012') "
VstarConn.Execute " Insert Into NewCaptions Values('02013','Invalid Password.','02013') "
VstarConn.Execute " Insert Into NewCaptions Values('02014','Unacceptable Password :: Please Type New Password','02014') "
VstarConn.Execute " Insert Into NewCaptions Values('02015','Password for','02015') "
VstarConn.Execute " Insert Into NewCaptions Values('02016',' Changed Successfully','02016') "
VstarConn.Execute " Insert Into NewCaptions Values('03001','Select Path for .DAT File','03001') "
VstarConn.Execute " Insert Into NewCaptions Values('03002','Select Directory','03002') "
VstarConn.Execute " Insert Into NewCaptions Values('03003','Drive:','03003') "
VstarConn.Execute " Insert Into NewCaptions Values('03004','Select','03004') "
VstarConn.Execute " Insert Into NewCaptions Values('03005','Cancel','03005') "
VstarConn.Execute " Insert Into NewCaptions Values('04001','Send Report to. . .','04001') "
VstarConn.Execute " Insert Into NewCaptions Values('04002','Type Subject here:','04002') "
VstarConn.Execute " Insert Into NewCaptions Values('04003',' Send Report to whom?','04003') "
VstarConn.Execute " Insert Into NewCaptions Values('04004','Send to Manager','04004') "
VstarConn.Execute " Insert Into NewCaptions Values('04005','Send to each Employee','04005') "
VstarConn.Execute " Insert Into NewCaptions Values('05001','About Visual Star(VSTAR)','05001') "
VstarConn.Execute " Insert Into NewCaptions Values('05002','Visual System for Time Attendance Recording','05002') "
VstarConn.Execute " Insert Into NewCaptions Values('05003','Version','05003') "
VstarConn.Execute " Insert Into NewCaptions Values('05004','Print Electronics Equipments Pvt Ltd','05004') "
VstarConn.Execute " Insert Into NewCaptions Values('05005','PEEPL','05005') "
VstarConn.Execute " Insert Into NewCaptions Values('05006','E-mail','05006') "
VstarConn.Execute " Insert Into NewCaptions Values('05007','Support @printelectronics.com','05007') "
VstarConn.Execute " Insert Into NewCaptions Values('05008','sales@printelectronics.com','05008') "
VstarConn.Execute " Insert Into NewCaptions Values('05009','Also Visit us at: www.printelectronics.com','05009') "
VstarConn.Execute " Insert Into NewCaptions Values('05010','Print Electronics Equipment Pvt Ltd','05010') "
VstarConn.Execute " Insert Into NewCaptions Values('06001','Print Electronics Administrator','06001') "
VstarConn.Execute " Insert Into NewCaptions Values('06002','Reset &' || 'Exclusive Lock','06002') "
VstarConn.Execute " Insert Into NewCaptions Values('06003','Reset &' || 'Daily Lock','06003') "
VstarConn.Execute " Insert Into NewCaptions Values('06004','Reset &' || 'Monthly Lock','06004') "
VstarConn.Execute " Insert Into NewCaptions Values('06005','Reset &' || 'Yearly Lock','06005') "
VstarConn.Execute " Insert Into NewCaptions Values('06006','Please Confirm that Daily Data is not being Processed elsewhere','06006') "
VstarConn.Execute " Insert Into NewCaptions Values('06007','Please Confirm that Monthly Data is not being Processed elsewhere','06007') "
VstarConn.Execute " Insert Into NewCaptions Values('06008','Please Confirm that Yearly Leaves are not being Processed elsewhere','06008') "
VstarConn.Execute " Insert Into NewCaptions Values('07001','Avail Entry','07001') "
VstarConn.Execute " Insert Into NewCaptions Values('07002','Employee Code','07002') "
VstarConn.Execute " Insert Into NewCaptions Values('07003','Name','07003') "
VstarConn.Execute " Insert Into NewCaptions Values('07004','Leave Information','07004') "
VstarConn.Execute " Insert Into NewCaptions Values('07005','Leave Code','07005') "
VstarConn.Execute " Insert Into NewCaptions Values('07006','Leave type','07006') "
VstarConn.Execute " Insert Into NewCaptions Values('07007','No of Days','07007') "
VstarConn.Execute " Insert Into NewCaptions Values('07008','Balance','07008') "
VstarConn.Execute " Insert Into NewCaptions Values('07009','Leave Code','07009') "
VstarConn.Execute " Insert Into NewCaptions Values('07010','Leave From','07010') "
VstarConn.Execute " Insert Into NewCaptions Values('07011','Leave To','07011') "
VstarConn.Execute " Insert Into NewCaptions Values('07012','Leave Days','07012') "
VstarConn.Execute " Insert Into NewCaptions Values('07013','Code','07013') "
VstarConn.Execute " Insert Into NewCaptions Values('07014','Name','07014') "
VstarConn.Execute " Insert Into NewCaptions Values('07015','Balance','07015') "
VstarConn.Execute " Insert Into NewCaptions Values('07016','Employee Not Found','07016') "
VstarConn.Execute " Insert Into NewCaptions Values('07017','Please Select the Leave to be Availed','07017') "
VstarConn.Execute " Insert Into NewCaptions Values('07018','Leaves cannot be Availed for 0 Number of days','07018') "
VstarConn.Execute " Insert Into NewCaptions Values('07019','Already Availed','07019') "
VstarConn.Execute " Insert Into NewCaptions Values('07020',' Times, Still Continue','07020') "
VstarConn.Execute " Insert Into NewCaptions Values('07021','Maximum','07021') "
VstarConn.Execute " Insert Into NewCaptions Values('07022',' Leave(s) Can be Availed ,Still Continue?','07022') "
VstarConn.Execute " Insert Into NewCaptions Values('07023','Minimum','07023') "
VstarConn.Execute " Insert Into NewCaptions Values('07024','No Leave Balances are Remaining','07024') "
VstarConn.Execute " Insert Into NewCaptions Values('07025','Still Continue ?','07025') "

VstarConn.Execute " Insert Into NewCaptions Values('07026','Balance is already Over, Still Avail Leaves','07026') "
VstarConn.Execute " Insert Into NewCaptions Values('07027','Leaves already Availed on the One of the above Selected Date(s)','07027') "
VstarConn.Execute " Insert Into NewCaptions Values('07028','This Employee is absent on Immediate days.','07028') "
VstarConn.Execute " Insert Into NewCaptions Values('07029',' Leave applied','07029') "
VstarConn.Execute " Insert Into NewCaptions Values('07030','.  Accept this leave?','07030') "
VstarConn.Execute " Insert Into NewCaptions Values('07031','The Employee has Availed Leaves for Continuous Days , Still Continue?','07031') "
VstarConn.Execute " Insert Into NewCaptions Values('07032','Monthly Transaction File not Found','07032') "
VstarConn.Execute " Insert Into NewCaptions Values('07033','Updation Cannot be Done','07033') "
VstarConn.Execute " Insert Into NewCaptions Values('07034',' for Leave Deletion','07034') "
VstarConn.Execute " Insert Into NewCaptions Values('07035','This Leave has been Deleted from the Master','07035') "
VstarConn.Execute " Insert Into NewCaptions Values('07036','This may cause this Form to function improperly','07036') "
VstarConn.Execute " Insert Into NewCaptions Values('07037','Operation Aborted','07037') "
VstarConn.Execute " Insert Into NewCaptions Values('07038','CO for extra work done on','07038') "
VstarConn.Execute " Insert Into NewCaptions Values('07039','Please enter CO Entry Date','07039') "
VstarConn.Execute " Insert Into NewCaptions Values('07040','From and To Date must be same','07040') "
VstarConn.Execute " Insert Into NewCaptions Values('07041','No CO found for specified Date','07041') "
VstarConn.Execute " Insert Into NewCaptions Values('07042','CO Availment Date Limit Over, Still Continue ?','07042') "
VstarConn.Execute " Insert Into NewCaptions Values('07043','CO cannot be availed more than the available balance.','07043') "
VstarConn.Execute " Insert Into NewCaptions Values('07044','Entry Date cannot be as same as From Date','07044') "
VstarConn.Execute " Insert Into NewCaptions Values('08001','BackUp and Restore','08001') "
VstarConn.Execute " Insert Into NewCaptions Values('08002','Operation','08002') "
VstarConn.Execute " Insert Into NewCaptions Values('08003','&' || 'BackUp','08003') "
VstarConn.Execute " Insert Into NewCaptions Values('08004','&' || 'Restore','08004') "
VstarConn.Execute " Insert Into NewCaptions Values('08005','Please Select the Type of Operation First','08005') "
VstarConn.Execute " Insert Into NewCaptions Values('08006','Please Select the Directory for Back Up','08006') "
VstarConn.Execute " Insert Into NewCaptions Values('08007','Please Select the File to Restore','08007') "
VstarConn.Execute " Insert Into NewCaptions Values('08008','Are You Sure to BackUp the Database','08008') "
VstarConn.Execute " Insert Into NewCaptions Values('08009','Are You Sure to Restore the Database','08009') "
VstarConn.Execute " Insert Into NewCaptions Values('08010','Error Restoring Connection','08010') "
VstarConn.Execute " Insert Into NewCaptions Values('08011','Recommended to Quit and Restart the Application','08011') "
VstarConn.Execute " Insert Into NewCaptions Values('09001','Select Date','09001') "
VstarConn.Execute " Insert Into NewCaptions Values('09002','Please Select the Date First','09002') "
VstarConn.Execute " Insert Into NewCaptions Values('10001','Category Master','10001') "
VstarConn.Execute " Insert Into NewCaptions Values('10002','Category Code','10002') "
VstarConn.Execute " Insert Into NewCaptions Values('10003','Category Name','10003') "
VstarConn.Execute " Insert Into NewCaptions Values('10004','Info','10004') "
VstarConn.Execute " Insert Into NewCaptions Values('10005','Late coming/Early going Rules','10005') "
VstarConn.Execute " Insert Into NewCaptions Values('10006','Allow Employee to come late by','10006') "
VstarConn.Execute " Insert Into NewCaptions Values('10007','Allow Employee to go early by','10007') "
VstarConn.Execute " Insert Into NewCaptions Values('10008','Ignore early arrival before shift by','10008') "
VstarConn.Execute " Insert Into NewCaptions Values('10009','Ignore late going after shift by','10009') "
VstarConn.Execute " Insert Into NewCaptions Values('10010','Cut half day if late coming by','10010') "
VstarConn.Execute " Insert Into NewCaptions Values('10011','Cut half day if early going by','10011') "
VstarConn.Execute " Insert Into NewCaptions Values('10012','Comp .Off Rule for Normal Days','10012') "
VstarConn.Execute " Insert Into NewCaptions Values('10013','Credit half day for working more than','10013') "
VstarConn.Execute " Insert Into NewCaptions Values('10014','Credit full day for working more than','10014') "
VstarConn.Execute " Insert Into NewCaptions Values('10015','Comp .Off Rule for other Days','10015') "
VstarConn.Execute " Insert Into NewCaptions Values('10016','Category Code cannot be blank','10016') "
VstarConn.Execute " Insert Into NewCaptions Values('10017','Category already exists','10017') "
VstarConn.Execute " Insert Into NewCaptions Values('10018','Category with the Same Name Already Exists','10018') "
VstarConn.Execute " Insert Into NewCaptions Values('10019','Category Name cannot be blank','10019') "
VstarConn.Execute " Insert Into NewCaptions Values('11001','Compact Database','11001') "
VstarConn.Execute " Insert Into NewCaptions Values('11002','Note :: This option will Compact the Database file.This will Save the Disk space where the Application file is installed and help in better Application Performance . Before running this option make sure the user has logged in Exclusively and no other use','11002') "
VstarConn.Execute " Insert Into NewCaptions Values('11003','Compact &' || 'Database','11003') "
VstarConn.Execute " Insert Into NewCaptions Values('11004','Make Sure you Read the Note on the Form before Running this Option.','11004') "
VstarConn.Execute " Insert Into NewCaptions Values('11005','Database Compacted','11005') "
VstarConn.Execute " Insert Into NewCaptions Values('12001','Company Master','12001') "
VstarConn.Execute " Insert Into NewCaptions Values('12002','Company Code','12002') "
VstarConn.Execute " Insert Into NewCaptions Values('12003','Company Name','12003') "
VstarConn.Execute " Insert Into NewCaptions Values('12004','Company not Found','12004') "
VstarConn.Execute " Insert Into NewCaptions Values('12005','Cannot Add More than','12005') "
VstarConn.Execute " Insert Into NewCaptions Values('12006',' Companies.','12006') "
VstarConn.Execute " Insert Into NewCaptions Values('12007','Company Code cannot be blank','12007') "
VstarConn.Execute " Insert Into NewCaptions Values('12008','Company Code Already Exists','12008') "
VstarConn.Execute " Insert Into NewCaptions Values('12009','Company Name cannot be blank','12009') "
VstarConn.Execute " Insert Into NewCaptions Values('13001','Correction','13001') "
VstarConn.Execute " Insert Into NewCaptions Values('13002','Employee Code','13002') "
VstarConn.Execute " Insert Into NewCaptions Values('13003','Name','13003') "
VstarConn.Execute " Insert Into NewCaptions Values('13004','Attendance Records','13004') "
VstarConn.Execute " Insert Into NewCaptions Values('13005','Attendance Details','13005') "
VstarConn.Execute " Insert Into NewCaptions Values('13006','Work Hrs','13006') "
VstarConn.Execute " Insert Into NewCaptions Values('13007','Present','13007') "
VstarConn.Execute " Insert Into NewCaptions Values('13008','Details','13008') "
VstarConn.Execute " Insert Into NewCaptions Values('13009','Misc.','13009') "
VstarConn.Execute " Insert Into NewCaptions Values('13010','Present Days','13010') "
VstarConn.Execute " Insert Into NewCaptions Values('13011','Rest Hrs','13011') "
VstarConn.Execute " Insert Into NewCaptions Values('13012','CO Days','13012') "
VstarConn.Execute " Insert Into NewCaptions Values('13013','Time','13013') "
VstarConn.Execute " Insert Into NewCaptions Values('13014','Irregular Entries','13014') "
VstarConn.Execute " Insert Into NewCaptions Values('13015','2nd','13015') "
VstarConn.Execute " Insert Into NewCaptions Values('13016','4th','13016') "
VstarConn.Execute " Insert Into NewCaptions Values('13017','6th','13017') "
VstarConn.Execute " Insert Into NewCaptions Values('13018','3rd','13018') "
VstarConn.Execute " Insert Into NewCaptions Values('13019','5th','13019') "
VstarConn.Execute " Insert Into NewCaptions Values('13020','7th','13020') "
VstarConn.Execute " Insert Into NewCaptions Values('13021','On Duty','13021') "
VstarConn.Execute " Insert Into NewCaptions Values('13022','Off Duty','13022') "
VstarConn.Execute " Insert Into NewCaptions Values('13023','Permission','13023') "
VstarConn.Execute " Insert Into NewCaptions Values('13024','Late Card','13024') "
VstarConn.Execute " Insert Into NewCaptions Values('13025','Early Card','13025') "
VstarConn.Execute " Insert Into NewCaptions Values('13026','&' || 'Shift','13026') "
VstarConn.Execute " Insert Into NewCaptions Values('13027','&' || 'Record','13027') "
VstarConn.Execute " Insert Into NewCaptions Values('13028','&' || 'Status','13028') "
VstarConn.Execute " Insert Into NewCaptions Values('13029','&' || 'On Duty','13029') "
VstarConn.Execute " Insert Into NewCaptions Values('13030','&' || 'Off Duty','13030') "
VstarConn.Execute " Insert Into NewCaptions Values('13031','&' || 'Time','13031') "
VstarConn.Execute " Insert Into NewCaptions Values('13032','OT/CO','13032') "
VstarConn.Execute " Insert Into NewCaptions Values('13033','&' || 'CO','13033') "
VstarConn.Execute " Insert Into NewCaptions Values('13034','This Employee does not have Overtime or Comp Off','13034') "
VstarConn.Execute " Insert Into NewCaptions Values('13035','File not found for the Month of','13035') "
VstarConn.Execute " Insert Into NewCaptions Values('13036','No Records Found For the Employee','13036') "
VstarConn.Execute " Insert Into NewCaptions Values('13037','Error Finding the Employee Record for the Date','13037') "
VstarConn.Execute " Insert Into NewCaptions Values('13038','Invalid Value','13038') "
VstarConn.Execute " Insert Into NewCaptions Values('13039','Please Select the Shift','13039') "
VstarConn.Execute " Insert Into NewCaptions Values('13040','Invalid Late Card Number','13040') "
VstarConn.Execute " Insert Into NewCaptions Values('13041','Invalid Early Card Number','13041') "
VstarConn.Execute " Insert Into NewCaptions Values('13042','Minutes Should be less than 60','13042') "
VstarConn.Execute " Insert Into NewCaptions Values('13043','Invalid value :: Cannot be Greater than 48','13043') "
VstarConn.Execute " Insert Into NewCaptions Values('13044','Departure Time Cannot be 0 if Arrival Time is Greater than 0','13044') "
VstarConn.Execute " Insert Into NewCaptions Values('13045','Arrival Time Cannot be Greater then Departure Time','13045') "
VstarConn.Execute " Insert Into NewCaptions Values('13046','Arrival Time Cannot be 0 if Departure Time is Greater than 0','13046') "
VstarConn.Execute " Insert Into NewCaptions Values('13047','Punch Time Should be between Arrival Time and Departure Time','13047') "
VstarConn.Execute " Insert Into NewCaptions Values('13048','To Time Should be Greater than From Time','13048') "
VstarConn.Execute " Insert Into NewCaptions Values('13049','On Duty From Time Cannot Be Greater than On Duty To Time','13049') "
VstarConn.Execute " Insert Into NewCaptions Values('13050','On Duty From punch Missing','13050') "
VstarConn.Execute " Insert Into NewCaptions Values('13051','On Duty From Time Should be between Arrival Time and Departure Time','13051') "
VstarConn.Execute " Insert Into NewCaptions Values('13052','On Duty To Time Should be between Arrival Time and Departure Time','13052') "
VstarConn.Execute " Insert Into NewCaptions Values('13053','Off Duty From punch Missing','13053') "
VstarConn.Execute " Insert Into NewCaptions Values('13054','Off Duty From Time Should be between Arrival Time and Departure Time','13054') "
VstarConn.Execute " Insert Into NewCaptions Values('13055','Off Duty To Time Should be between Arrival Time and Departure Time','13055') "
VstarConn.Execute " Insert Into NewCaptions Values('13056','2nd punch Missing','13056') "
VstarConn.Execute " Insert Into NewCaptions Values('13057','3rd punch Missing','13057') "
VstarConn.Execute " Insert Into NewCaptions Values('13058','4th punch Missing','13058') "
VstarConn.Execute " Insert Into NewCaptions Values('13059','5th punch Missing','13059') "
VstarConn.Execute " Insert Into NewCaptions Values('13060','5th punch Missing','13060') "
VstarConn.Execute " Insert Into NewCaptions Values('13061','2nd punch cannot be Greater than 3rd Punch','13061') "
VstarConn.Execute " Insert Into NewCaptions Values('13062','3rd punch cannot be Greater than 4th Punch','13062') "
VstarConn.Execute " Insert Into NewCaptions Values('13063','4th punch cannot be Greater than 5th Punch','13063') "
VstarConn.Execute " Insert Into NewCaptions Values('13064','5th punch cannot be Greater than 6th Punch','13064') "
VstarConn.Execute " Insert Into NewCaptions Values('13065','6th punch cannot be Greater than 7th Punch','13065') "
VstarConn.Execute " Insert Into NewCaptions Values('13066','CO not Found :: Leave Balance File for the Current Year not Updated','13066') "
VstarConn.Execute " Insert Into NewCaptions Values('13067','CO not Found in Leave Master','13067') "
VstarConn.Execute " Insert Into NewCaptions Values('13068','Leave Balanace File for the Current Year not Found','13068') "
VstarConn.Execute " Insert Into NewCaptions Values('13069','Please Create it First and then do the Daily Process','13069') "
VstarConn.Execute " Insert Into NewCaptions Values('13070','NO CO Rule is set','13070') "
VstarConn.Execute " Insert Into NewCaptions Values('14001','Change Period','14001') "
VstarConn.Execute " Insert Into NewCaptions Values('14002','From Day','14002') "
VstarConn.Execute " Insert Into NewCaptions Values('14003','To Day','14003') "
VstarConn.Execute " Insert Into NewCaptions Values('14004','&' || 'Overwtite Week Off''s','14004') "
VstarConn.Execute " Insert Into NewCaptions Values('14005','Overwrite &' || 'Holidays','14005') "
VstarConn.Execute " Insert Into NewCaptions Values('14006','&' || 'Change','14006') "
VstarConn.Execute " Insert Into NewCaptions Values('14007','Periodic Shift Updation done','14007') "
VstarConn.Execute " Insert Into NewCaptions Values('14008','Please Select the Month First','14008') "
VstarConn.Execute " Insert Into NewCaptions Values('14009','Please Select the Year First','14009') "
VstarConn.Execute " Insert Into NewCaptions Values('14010','Shift File not found for the Month of','14010') "
VstarConn.Execute " Insert Into NewCaptions Values('14011','Please Create it First Using Shift Creation','14011') "
VstarConn.Execute " Insert Into NewCaptions Values('14012','Please Select the Day from the where Shifts are to be Updated','14012') "
VstarConn.Execute " Insert Into NewCaptions Values('14013','Please Select the Day Till the where Shifts are to be Updated','14013') "
VstarConn.Execute " Insert Into NewCaptions Values('14014','From Period cannot be Greater than To Period','14014') "
VstarConn.Execute " Insert Into NewCaptions Values('14015','Please Select the Shift','14015') "
VstarConn.Execute " Insert Into NewCaptions Values('14016','Please Select the Employee First','14016') "
VstarConn.Execute " Insert Into NewCaptions Values('15001','Credit Entry','15001') "
VstarConn.Execute " Insert Into NewCaptions Values('15002','Credit On','15002') "
VstarConn.Execute " Insert Into NewCaptions Values('15003','Please Select the Leave to be Credited','15003') "
VstarConn.Execute " Insert Into NewCaptions Values('15004','Leaves cannot be Credited for 0 Number of days','15004') "
VstarConn.Execute " Insert Into NewCaptions Values('15005','Days To be Credited Must be Divisible by 0.50','15005') "
VstarConn.Execute " Insert Into NewCaptions Values('15006','From date Cannot be Empty','15006') "
VstarConn.Execute " Insert Into NewCaptions Values('15007','Leaves already Credited on the above Selected Date','15007') "
VstarConn.Execute " Insert Into NewCaptions Values('15008','Maximum Credit every year are','15008') "
VstarConn.Execute " Insert Into NewCaptions Values('15009',' days','15009') "
VstarConn.Execute " Insert Into NewCaptions Values('15010','Credit All days?','15010') "
VstarConn.Execute " Insert Into NewCaptions Values('16001','Customize leave Codes','16001') "
VstarConn.Execute " Insert Into NewCaptions Values('16002','Keep default Codes for present /Absent/Week Off /Holiday','16002') "
VstarConn.Execute " Insert Into NewCaptions Values('16003','Change Leave Codes for present /Absent/Week Off /Holiday','16003') "
VstarConn.Execute " Insert Into NewCaptions Values('16004','Type','16004') "
VstarConn.Execute " Insert Into NewCaptions Values('16005','Existing Codes','16005') "
VstarConn.Execute " Insert Into NewCaptions Values('16006','New Codes','16006') "
VstarConn.Execute " Insert Into NewCaptions Values('16007','Absent Days','16007') "
VstarConn.Execute " Insert Into NewCaptions Values('16008','Present Days','16008') "
VstarConn.Execute " Insert Into NewCaptions Values('16009','Week Offs','16009') "
VstarConn.Execute " Insert Into NewCaptions Values('16010','Holidays','16010') "
VstarConn.Execute " Insert Into NewCaptions Values('16011','&' || 'Save and Exit','16011') "
VstarConn.Execute " Insert Into NewCaptions Values('16012','Please Enter the Absent Code','16012') "
VstarConn.Execute " Insert Into NewCaptions Values('16013','Please Enter the Present Code','16013') "
VstarConn.Execute " Insert Into NewCaptions Values('16014','Please Enter the Week Off Code','16014') "
VstarConn.Execute " Insert Into NewCaptions Values('16015','Please Enter the Holidays Code','16015') "
VstarConn.Execute " Insert Into NewCaptions Values('16016','Duplicate codes not Allowed','16016') "
VstarConn.Execute " Insert Into NewCaptions Values('16017','Are You Sure to Change the Custom Codes','16017') "
VstarConn.Execute " Insert Into NewCaptions Values('16018','Error in Updating Custom Codes','16018') "
VstarConn.Execute " Insert Into NewCaptions Values('16019','Please Create Yearly Leave Files','16019') "
VstarConn.Execute " Insert Into NewCaptions Values('17001','Daily Processing','17001') "
VstarConn.Execute " Insert Into NewCaptions Values('17002','Processing Dates','17002') "
VstarConn.Execute " Insert Into NewCaptions Values('17003','&' || 'From Date','17003') "
VstarConn.Execute " Insert Into NewCaptions Values('17004','&' || 'To Date','17004') "
VstarConn.Execute " Insert Into NewCaptions Values('17005','Select Dat File','17005') "
VstarConn.Execute " Insert Into NewCaptions Values('17006','&' || 'Exclude Dat Files','17006') "
VstarConn.Execute " Insert Into NewCaptions Values('17007','Retreiving Records from the Dat File ..','17007') "
VstarConn.Execute " Insert Into NewCaptions Values('17008','Processing Records :: Please Wait ..','17008') "
VstarConn.Execute " Insert Into NewCaptions Values('17009','&' || 'Process','17009') "
VstarConn.Execute " Insert Into NewCaptions Values('17010','This Will Clear your All Dat Files Selection','17010') "
VstarConn.Execute " Insert Into NewCaptions Values('17011','Daily Process is Aborted','17011') "
VstarConn.Execute " Insert Into NewCaptions Values('17012','Daliy Process is Over','17012') "
VstarConn.Execute " Insert Into NewCaptions Values('17013','Software Locked :: Cannot Process','17013') "
VstarConn.Execute " Insert Into NewCaptions Values('17014','Contact Print Electronics','17014') "
VstarConn.Execute " Insert Into NewCaptions Values('17015','Duplicate File Names not Allowed','17015') "
VstarConn.Execute " Insert Into NewCaptions Values('17016','Since all the Shift Files Necessary for Processing are not Created ,Processing Cannot Continue.','17016') "
VstarConn.Execute " Insert Into NewCaptions Values('17017','Shift file for the month of','17017') "
VstarConn.Execute " Insert Into NewCaptions Values('17018',' not available','17018') "
VstarConn.Execute " Insert Into NewCaptions Values('17019','Do you want to create it','17019') "
VstarConn.Execute " Insert Into NewCaptions Values('17020','Please do the Processing for','17020') "

VstarConn.Execute " Insert Into NewCaptions Values('17021','Remove','17021') "
VstarConn.Execute " Insert Into NewCaptions Values('17022','Click to Toggle Selection','17022') "
VstarConn.Execute " Insert Into NewCaptions Values('18001','Select Source','18001') "
VstarConn.Execute " Insert Into NewCaptions Values('18002','&' || 'Files','18002') "
VstarConn.Execute " Insert Into NewCaptions Values('18003','&' || 'Drives','18003') "
VstarConn.Execute " Insert Into NewCaptions Values('18004','Fold&' || 'ers','18004') "
VstarConn.Execute " Insert Into NewCaptions Values('19001','Declare Holiday/ WeekOff','19001') "
VstarConn.Execute " Insert Into NewCaptions Values('19002','Compensate On','19002') "
VstarConn.Execute " Insert Into NewCaptions Values('19003','As','19003') "
VstarConn.Execute " Insert Into NewCaptions Values('19004','Add this Holiday/WeekOff for all categories','19004') "
VstarConn.Execute " Insert Into NewCaptions Values('19005','Compensate date','19005') "
VstarConn.Execute " Insert Into NewCaptions Values('19006','Declare as','19006') "
VstarConn.Execute " Insert Into NewCaptions Values('19007','WeekOff','19007') "
VstarConn.Execute " Insert Into NewCaptions Values('19008','Holiday','19008') "
VstarConn.Execute " Insert Into NewCaptions Values('19009','Category Does not Exist :: Cannot Display the Record','19009') "
VstarConn.Execute " Insert Into NewCaptions Values('19010','Category cannot be blank','19010') "
VstarConn.Execute " Insert Into NewCaptions Values('19011','Blank Category Master  :: Cannot Add the Record','19011') "
VstarConn.Execute " Insert Into NewCaptions Values('19012','Date cannot be blank','19012') "
VstarConn.Execute " Insert Into NewCaptions Values('19013','Holiday/Week Off Date','19013') "
VstarConn.Execute " Insert Into NewCaptions Values('19014','Compensate Date','19014') "
VstarConn.Execute " Insert Into NewCaptions Values('19015','Holiday Date and Compensate Date cannot be Same','19015') "
VstarConn.Execute " Insert Into NewCaptions Values('19016','Description cannot be Blank','19016') "
VstarConn.Execute " Insert Into NewCaptions Values('19017','It''s a Week Off,Cannot Declare Holiday/Week Off on the Same Date','19017') "
VstarConn.Execute " Insert Into NewCaptions Values('19018','Holiday Already Declared on the Selected Date','19018') "
VstarConn.Execute " Insert Into NewCaptions Values('19019','No Employees :: Cannot Add Holidays.','19019') "
VstarConn.Execute " Insert Into NewCaptions Values('20001','DEPARTMENT MASTER','20001') "
VstarConn.Execute " Insert Into NewCaptions Values('20002','Strength','20002') "
VstarConn.Execute " Insert Into NewCaptions Values('20003','Department not Found','20003') "
VstarConn.Execute " Insert Into NewCaptions Values('20004','Department Code cannot be blank','20004') "
VstarConn.Execute " Insert Into NewCaptions Values('20005','Department Code Already Exists','20005') "
VstarConn.Execute " Insert Into NewCaptions Values('20006','Department Name cannot be blank','20006') "
VstarConn.Execute " Insert Into NewCaptions Values('20007','Department with Same Code Already Exists','20007') "
VstarConn.Execute " Insert Into NewCaptions Values('21001','Make DSN','21001') "
VstarConn.Execute " Insert Into NewCaptions Values('21002','&' || 'Back End','21002') "
VstarConn.Execute " Insert Into NewCaptions Values('21003','&' || 'User Name','21003') "
VstarConn.Execute " Insert Into NewCaptions Values('21004','&' || 'Password','21004') "
VstarConn.Execute " Insert Into NewCaptions Values('21005','DSN Name','21005') "
VstarConn.Execute " Insert Into NewCaptions Values('21006','Server &' || 'Name','21006') "
VstarConn.Execute " Insert Into NewCaptions Values('21007','P&' || 'ath','21007') "
VstarConn.Execute " Insert Into NewCaptions Values('21008','Due to Some Reasons the DSN may be Corrupted or Deleted. DSN can be Created with a Valid User Name, Password (Case Sensitive) and a Server Name. Please Contact Your System Administrator or Print Electronics for further Details.','21008') "
VstarConn.Execute " Insert Into NewCaptions Values('21009','Due to Some Reasons the DSN may be Corrupted or Deleted. DSN can be Created with a Valid Password(Case Sensitive), Also Enter a Valid MDB File Path. Please Contact Your System Administrator or Print Electronics for further Details.','21009') "
VstarConn.Execute " Insert Into NewCaptions Values('21010','Details not Yet Available Yet. Please Contact Print Electronics for further Details.','21010') "
VstarConn.Execute " Insert Into NewCaptions Values('21011','&' || 'Show System DSN Wizard','21011') "
VstarConn.Execute " Insert Into NewCaptions Values('21012','Please Enter Server Name','21012') "
VstarConn.Execute " Insert Into NewCaptions Values('21013','Please Enter UserName','21013') "
VstarConn.Execute " Insert Into NewCaptions Values('21014','Please Enter Password','21014') "
VstarConn.Execute " Insert Into NewCaptions Values('21015','Please Enter Database Path','21015') "
VstarConn.Execute " Insert Into NewCaptions Values('21016','DSN Created Successfully','21016') "
VstarConn.Execute " Insert Into NewCaptions Values('22001','Edit Paid Days','22001') "
VstarConn.Execute " Insert Into NewCaptions Values('22002','Employee Code','22002') "
VstarConn.Execute " Insert Into NewCaptions Values('22003','Paid Days','22003') "
VstarConn.Execute " Insert Into NewCaptions Values('22004','Present','22004') "
VstarConn.Execute " Insert Into NewCaptions Values('22005','Absent','22005') "
VstarConn.Execute " Insert Into NewCaptions Values('22006','WeekOff','22006') "
VstarConn.Execute " Insert Into NewCaptions Values('22007','Holiday','22007') "
VstarConn.Execute " Insert Into NewCaptions Values('22008','Please Create it First','22008') "
VstarConn.Execute " Insert Into NewCaptions Values('23001','Employee Master','23001') "
VstarConn.Execute " Insert Into NewCaptions Values('23002','Find Employee with Employee code','23002') "
VstarConn.Execute " Insert Into NewCaptions Values('23003','or having name','23003') "
VstarConn.Execute " Insert Into NewCaptions Values('23004','Official Details','23004') "
VstarConn.Execute " Insert Into NewCaptions Values('23005','Personal Details','23005') "
VstarConn.Execute " Insert Into NewCaptions Values('23006','Other Details','23006') "
VstarConn.Execute " Insert Into NewCaptions Values('23007','Emp Code','23007') "
VstarConn.Execute " Insert Into NewCaptions Values('23008','Employee Name','23008') "
VstarConn.Execute " Insert Into NewCaptions Values('23009','Emp Card','23009') "
VstarConn.Execute " Insert Into NewCaptions Values('23010','Join Date','23010') "
VstarConn.Execute " Insert Into NewCaptions Values('23011','Conf. Date','23011') "
VstarConn.Execute " Insert Into NewCaptions Values('23012','Code No','23012') "
VstarConn.Execute " Insert Into NewCaptions Values('23013','Card No','23013') "
VstarConn.Execute " Insert Into NewCaptions Values('23014','Designation','23014') "
VstarConn.Execute " Insert Into NewCaptions Values('23015','Identification','23015') "
VstarConn.Execute " Insert Into NewCaptions Values('23016','Min. Entry','23016') "
VstarConn.Execute " Insert Into NewCaptions Values('23017','Compensatory Off','23017') "
VstarConn.Execute " Insert Into NewCaptions Values('23018','Autoshift Change','23018') "
VstarConn.Execute " Insert Into NewCaptions Values('23019','Travel By','23019') "
VstarConn.Execute " Insert Into NewCaptions Values('23020','Division','23020') "
VstarConn.Execute " Insert Into NewCaptions Values('23021','Working Schedule','23021') "
VstarConn.Execute " Insert Into NewCaptions Values('23022','Define Schedule','23022') "
VstarConn.Execute " Insert Into NewCaptions Values('23023','Past Employee','23023') "
VstarConn.Execute " Insert Into NewCaptions Values('23024','Left Date','23024') "
VstarConn.Execute " Insert Into NewCaptions Values('23025','Date of Birth','23025') "
VstarConn.Execute " Insert Into NewCaptions Values('23026','Blood Group','23026') "
VstarConn.Execute " Insert Into NewCaptions Values('23027','Date of join','23027') "
VstarConn.Execute " Insert Into NewCaptions Values('23028','Confirm Date','23028') "
VstarConn.Execute " Insert Into NewCaptions Values('23029','Sex','23029') "
VstarConn.Execute " Insert Into NewCaptions Values('23030','E-Mail ID','23030') "
VstarConn.Execute " Insert Into NewCaptions Values('23031','Basic Salary','23031') "
VstarConn.Execute " Insert Into NewCaptions Values('23032','Reference','23032') "
VstarConn.Execute " Insert Into NewCaptions Values('23033','Address','23033') "
VstarConn.Execute " Insert Into NewCaptions Values('23034','City','23034') "
VstarConn.Execute " Insert Into NewCaptions Values('23035','Pin Code','23035') "
VstarConn.Execute " Insert Into NewCaptions Values('23036','Phone No','23036') "
VstarConn.Execute " Insert Into NewCaptions Values('23037','Permanent Address','23037') "
VstarConn.Execute " Insert Into NewCaptions Values('23038','HouseNo/Name','23038') "
VstarConn.Execute " Insert Into NewCaptions Values('23039','City/Village','23039') "
VstarConn.Execute " Insert Into NewCaptions Values('23040','District','23040') "
VstarConn.Execute " Insert Into NewCaptions Values('23041','Tel.No','23041') "
VstarConn.Execute " Insert Into NewCaptions Values('23042','Area','23042') "
VstarConn.Execute " Insert Into NewCaptions Values('23043','Road','23043') "
VstarConn.Execute " Insert Into NewCaptions Values('23044','State','23044') "
VstarConn.Execute " Insert Into NewCaptions Values('23045','Nationality','23045') "
VstarConn.Execute " Insert Into NewCaptions Values('23046','Special Comments','23046') "
VstarConn.Execute " Insert Into NewCaptions Values('23047','Record not Found','23047') "
VstarConn.Execute " Insert Into NewCaptions Values('23048','Please Enter Employee Code','23048') "
VstarConn.Execute " Insert Into NewCaptions Values('23049','Yearly Leave Files are Not Created :: Please Create Them','23049') "
VstarConn.Execute " Insert Into NewCaptions Values('23050','Maximum Employee(s) Allowed :','23050') "
VstarConn.Execute " Insert Into NewCaptions Values('23051','Employee Already Exists','23051') "
VstarConn.Execute " Insert Into NewCaptions Values('23052','Employee Card Number Should be of','23052') "
VstarConn.Execute " Insert Into NewCaptions Values('23053',' Characters','23053') "
VstarConn.Execute " Insert Into NewCaptions Values('23054','Card Number Already Exists','23054') "
VstarConn.Execute " Insert Into NewCaptions Values('23055','Please Enter Employee Name','23055') "
VstarConn.Execute " Insert Into NewCaptions Values('23056','Please Select Employee Category','23056') "
VstarConn.Execute " Insert Into NewCaptions Values('23057','Please Select Employee Department','23057') "
VstarConn.Execute " Insert Into NewCaptions Values('23058','Please Select Employee Group','23058') "
VstarConn.Execute " Insert Into NewCaptions Values('23059','Please Select Company Code','23059') "
VstarConn.Execute " Insert Into NewCaptions Values('23060','Please Define Employee Shift','23060') "
VstarConn.Execute " Insert Into NewCaptions Values('23061','Please Enter Employee Joindate','23061') "
VstarConn.Execute " Insert Into NewCaptions Values('23062','Employee Joindate Must be Less than Employee Shift Date','23062') "
VstarConn.Execute " Insert Into NewCaptions Values('23063','Employee Joindate Must be Less than Employee Leave Date','23063') "
VstarConn.Execute " Insert Into NewCaptions Values('23064','Employee Shift Date Must be Less than Employee Leave Date','23064') "
VstarConn.Execute " Insert Into NewCaptions Values('23065','Employee Birth Date Must be Less than Employee Join Date','23065') "
VstarConn.Execute " Insert Into NewCaptions Values('23066','Employee Confirm Date Must be Greater than Employee Join Date','23066') "
VstarConn.Execute " Insert Into NewCaptions Values('23067',' is a reserved Permission Card No.','23067') "
VstarConn.Execute " Insert Into NewCaptions Values('23068','Please Select OT Rule','23068') "
VstarConn.Execute " Insert Into NewCaptions Values('23069','Please Select CO Rule','23069') "
VstarConn.Execute " Insert Into NewCaptions Values('23070','Please Select Location Code','23070') "
VstarConn.Execute " Insert Into NewCaptions Values('24001','Schedule Master','24001') "
VstarConn.Execute " Insert Into NewCaptions Values('24002','General','24002') "
VstarConn.Execute " Insert Into NewCaptions Values('24003','Shift Info','24003') "
VstarConn.Execute " Insert Into NewCaptions Values('24004','Shift Type','24004') "
VstarConn.Execute " Insert Into NewCaptions Values('24005','Starting Date','24005') "
VstarConn.Execute " Insert Into NewCaptions Values('24006','Rotation Code','24006') "
VstarConn.Execute " Insert Into NewCaptions Values('24007','Shift Code','24007') "
VstarConn.Execute " Insert Into NewCaptions Values('24008','WeekOff','24008') "
VstarConn.Execute " Insert Into NewCaptions Values('24009','There is a weekOff on every','24009') "
VstarConn.Execute " Insert Into NewCaptions Values('24010','Of a week','24010') "
VstarConn.Execute " Insert Into NewCaptions Values('24011','Additional Week-Offs','24011') "
VstarConn.Execute " Insert Into NewCaptions Values('24012','There is a week Off every','24012') "
VstarConn.Execute " Insert Into NewCaptions Values('24013','There is a week Off on the first &' || '&' || ' third','24013') "
VstarConn.Execute " Insert Into NewCaptions Values('24014','There is a week Off on the second &' || '&' || ' fourth','24014') "
VstarConn.Execute " Insert Into NewCaptions Values('24015','Shift Date Cannot be Empty','24015') "
VstarConn.Execute " Insert Into NewCaptions Values('24016','ShifDate Cannot be Less then the Join date','24016') "
VstarConn.Execute " Insert Into NewCaptions Values('24017','Please Select the Type of Shift','24017') "
VstarConn.Execute " Insert Into NewCaptions Values('24018','Please Select the Shift','24018') "
VstarConn.Execute " Insert Into NewCaptions Values('24019','Please Select the Rotational Shift','24019') "
VstarConn.Execute " Insert Into NewCaptions Values('24020','Please Select the Week Off Before Selecting the Additional week Off','24020') "
VstarConn.Execute " Insert Into NewCaptions Values('24021','Please Select the Additional Week Off','24021') "
VstarConn.Execute " Insert Into NewCaptions Values('24022','Please Select the First and Third Week Off','24022') "
VstarConn.Execute " Insert Into NewCaptions Values('24023','Please Select the Second and Fourth Week Off','24023') "
VstarConn.Execute " Insert Into NewCaptions Values('24024','Details regarding Daily Processing','24024') "
VstarConn.Execute " Insert Into NewCaptions Values('24025','On Weekoff / Holiday do the following','24025') "
VstarConn.Execute " Insert Into NewCaptions Values('24026','Assign Previous day Shift (Transaction)','24026') "
VstarConn.Execute " Insert Into NewCaptions Values('24027','Assign Next day Shift (Schedule)','24027') "
VstarConn.Execute " Insert Into NewCaptions Values('24028','Assign the following Shift','24028') "
VstarConn.Execute " Insert Into NewCaptions Values('24029','Assign Auto shift if punch found','24029') "
VstarConn.Execute " Insert Into NewCaptions Values('24030','If Blank Shift found','24030') "
VstarConn.Execute " Insert Into NewCaptions Values('24031','Keep it blank','24031') "
VstarConn.Execute " Insert Into NewCaptions Values('24032','Assign this Shift','24032') "
VstarConn.Execute " Insert Into NewCaptions Values('24033','Please Select the Shift to be assigned for Week Off / Holiday','24033') "
VstarConn.Execute " Insert Into NewCaptions Values('24034','Please Select the Shift to be assigned if  Blank Shift is Found','24034') "
VstarConn.Execute " Insert Into NewCaptions Values('24035','    &' || 'Set for more employees','24035') "
VstarConn.Execute " Insert Into NewCaptions Values('25001','Encash Entry','25001') "
VstarConn.Execute " Insert Into NewCaptions Values('25002','Encash on','25002') "
VstarConn.Execute " Insert Into NewCaptions Values('25003','Please Select the Leave to be Encashed','25003') "
VstarConn.Execute " Insert Into NewCaptions Values('25004','Leaves cannot be Encashed for 0 Number of days','25004') "
VstarConn.Execute " Insert Into NewCaptions Values('25005','Days To be Encashed Must be Divisible by 0.50','25005') "
VstarConn.Execute " Insert Into NewCaptions Values('25006','Balance is already Over, Still Encash Leaves','25006') "
VstarConn.Execute " Insert Into NewCaptions Values('25007','Leaves already Encashed on the above Selected Date','25007') "
VstarConn.Execute " Insert Into NewCaptions Values('26001','Select Dat File','26001') "
VstarConn.Execute " Insert Into NewCaptions Values('26002','Remove','26002') "
VstarConn.Execute " Insert Into NewCaptions Values('26003','Daily Processing Done for','26003') "
VstarConn.Execute " Insert Into NewCaptions Values('26004','No Valid Entries Found in the Dat File(s) for','26004') "
VstarConn.Execute " Insert Into NewCaptions Values('28001','Administrative User','28001') "
VstarConn.Execute " Insert Into NewCaptions Values('28002','This User Will Have All the Administrative Rights','28002') "
VstarConn.Execute " Insert Into NewCaptions Values('28003','&' || 'User Name','28003') "
VstarConn.Execute " Insert Into NewCaptions Values('28004','&' || 'Password','28004') "
VstarConn.Execute " Insert Into NewCaptions Values('28005','&' || 'Re-Type Password','28005') "
VstarConn.Execute " Insert Into NewCaptions Values('28006','The Date Format Selected will Effect the Software throughout it''s Life Time.','28006') "
VstarConn.Execute " Insert Into NewCaptions Values('28007','&' || 'British e.g. 29/03/01','28007') "
VstarConn.Execute " Insert Into NewCaptions Values('28008','&' || 'American e.g. 03/29/01','28008') "
VstarConn.Execute " Insert Into NewCaptions Values('28009','Date Format','28009') "
VstarConn.Execute " Insert Into NewCaptions Values('28010','The Date Format You Have Selected is British','28010') "
VstarConn.Execute " Insert Into NewCaptions Values('28011','Are you Sure to Continue ?','28011') "
VstarConn.Execute " Insert Into NewCaptions Values('28012','The Date Format You Have Selected is American','28012') "
VstarConn.Execute " Insert Into NewCaptions Values('28013','The Passwords don''t Match Please Re-Enter','28013') "
VstarConn.Execute " Insert Into NewCaptions Values('28014','Encryption Error :: Try Another Password','28014') "
VstarConn.Execute " Insert Into NewCaptions Values('28015',' Added as Admin Successfully','28015') "
VstarConn.Execute " Insert Into NewCaptions Values('29001','Group Master','29001') "
VstarConn.Execute " Insert Into NewCaptions Values('29002','Group Code cannot be Blank','29002') "
VstarConn.Execute " Insert Into NewCaptions Values('29003','Group Already Exists','29003') "
VstarConn.Execute " Insert Into NewCaptions Values('29004','Group Description cannot be Blank','29004') "
VstarConn.Execute " Insert Into NewCaptions Values('30001','Holiday Master','30001') "
VstarConn.Execute " Insert Into NewCaptions Values('30002','Holiday Date','30002') "
VstarConn.Execute " Insert Into NewCaptions Values('30003','Name of Holiday','30003') "
VstarConn.Execute " Insert Into NewCaptions Values('30004','Add this holiday for all categories','30004') "
VstarConn.Execute " Insert Into NewCaptions Values('30005','Specific Category','30005') "
VstarConn.Execute " Insert Into NewCaptions Values('30006','Category cannot be blank','30006') "
VstarConn.Execute " Insert Into NewCaptions Values('30007','Blank Category Master  :: Cannot Add the Record','30007') "
VstarConn.Execute " Insert Into NewCaptions Values('30008','Date cannot be blank','30008') "
VstarConn.Execute " Insert Into NewCaptions Values('30009','Description cannot be Blank','30009') "

VstarConn.Execute " Insert Into NewCaptions Values('30010','Holiday Already mentioned for this category','30010') "
VstarConn.Execute " Insert Into NewCaptions Values('30011','Category Does not Exist :: Cannot Display the Record','30011') "
VstarConn.Execute " Insert Into NewCaptions Values('31001','Leave Master','31001') "
VstarConn.Execute " Insert Into NewCaptions Values('31002','Custom Codes','31002') "
VstarConn.Execute " Insert Into NewCaptions Values('31003','Leave Code','31003') "
VstarConn.Execute " Insert Into NewCaptions Values('31004','Name of Leave','31004') "
VstarConn.Execute " Insert Into NewCaptions Values('31005','Leave Balance','31005') "
VstarConn.Execute " Insert Into NewCaptions Values('31006','Definition','31006') "
VstarConn.Execute " Insert Into NewCaptions Values('31007','Count this leave in payable days','31007') "
VstarConn.Execute " Insert Into NewCaptions Values('31008','Keep balance for this leave','31008') "
VstarConn.Execute " Insert Into NewCaptions Values('31009','For the current year','31009') "
VstarConn.Execute " Insert Into NewCaptions Values('31010','Credit','31010') "
VstarConn.Execute " Insert Into NewCaptions Values('31011','days leaves for consumption','31011') "
VstarConn.Execute " Insert Into NewCaptions Values('31012','Allow maximum','31012') "
VstarConn.Execute " Insert Into NewCaptions Values('31013','days leave to be accumulated','31013') "
VstarConn.Execute " Insert Into NewCaptions Values('31014','Crediting for new employees','31014') "
VstarConn.Execute " Insert Into NewCaptions Values('31015','Credit immediately','31015') "
VstarConn.Execute " Insert Into NewCaptions Values('31016','Credit next year','31016') "
VstarConn.Execute " Insert Into NewCaptions Values('31017','While crediting','31017') "
VstarConn.Execute " Insert Into NewCaptions Values('31018','Credit leaves full','31018') "
VstarConn.Execute " Insert Into NewCaptions Values('31019','Credit in proportion','31019') "
VstarConn.Execute " Insert Into NewCaptions Values('31020','All Categories','31020') "
VstarConn.Execute " Insert Into NewCaptions Values('31021','Specific','31021') "
VstarConn.Execute " Insert Into NewCaptions Values('31022','Mark Leaves','31022') "
VstarConn.Execute " Insert Into NewCaptions Values('31023','Including weekOff/Holidays','31023') "
VstarConn.Execute " Insert Into NewCaptions Values('31024','Excluding weekOff/Holidays','31024') "
VstarConn.Execute " Insert Into NewCaptions Values('31025','Decide while entering leave','31025') "
VstarConn.Execute " Insert Into NewCaptions Values('31026','At the End of the Year','31026') "
VstarConn.Execute " Insert Into NewCaptions Values('31027','Carry forward balance leaves','31027') "
VstarConn.Execute " Insert Into NewCaptions Values('31028','Encash balance leaves','31028') "
VstarConn.Execute " Insert Into NewCaptions Values('31029','Check following rules','31029') "
VstarConn.Execute " Insert Into NewCaptions Values('31030','Allow','31030') "
VstarConn.Execute " Insert Into NewCaptions Values('31031','times in ayear','31031') "
VstarConn.Execute " Insert Into NewCaptions Values('31032','Maximum','31032') "
VstarConn.Execute " Insert Into NewCaptions Values('31033','days at a time','31033') "
VstarConn.Execute " Insert Into NewCaptions Values('31034','Minimum','31034') "
VstarConn.Execute " Insert Into NewCaptions Values('31035','This Leave codes Defined Once will effect the Entire Software.','31035') "
VstarConn.Execute " Insert Into NewCaptions Values('31036','You Cannot Change them again or cannot make them Default.','31036') "
VstarConn.Execute " Insert Into NewCaptions Values('31037','Proceed further Y/N ?','31037') "
VstarConn.Execute " Insert Into NewCaptions Values('31038','Leave Code should be of atleast 2 Characters','31038') "
VstarConn.Execute " Insert Into NewCaptions Values('31039','Leave Code must not be Equal to ABSENT,PRESENT,WEEK OFF or HOLIDAY Code','31039') "
VstarConn.Execute " Insert Into NewCaptions Values('31040','Please Add Categories First','31040') "
VstarConn.Execute " Insert Into NewCaptions Values('31041','Please Select the Category First','31041') "
VstarConn.Execute " Insert Into NewCaptions Values('31042','Leave Already Defined','31042') "
VstarConn.Execute " Insert Into NewCaptions Values('31043','Leave Name cannot be Blank','31043') "
VstarConn.Execute " Insert Into NewCaptions Values('31044','Leave Acculumation should be Greater than Leave Credited','31044') "
VstarConn.Execute " Insert Into NewCaptions Values('31045','Please Select Full or Proportionate Leave','31045') "
VstarConn.Execute " Insert Into NewCaptions Values('31046','This may effect other Leave Specific Details throughout the Application','31046') "
VstarConn.Execute " Insert Into NewCaptions Values('32001','Lost Entry','32001') "
VstarConn.Execute " Insert Into NewCaptions Values('32002','Date of Punch','32002') "
VstarConn.Execute " Insert Into NewCaptions Values('32003','Time of Punch','32003') "
VstarConn.Execute " Insert Into NewCaptions Values('32004','Employee Name','32004') "
VstarConn.Execute " Insert Into NewCaptions Values('32005','Employee Cannot be blank','32005') "
VstarConn.Execute " Insert Into NewCaptions Values('32006','Time can''t be zero','32006') "
VstarConn.Execute " Insert Into NewCaptions Values('33001','Yearly Leave File Updation','33001') "
VstarConn.Execute " Insert Into NewCaptions Values('33002','Instructions','33002') "
VstarConn.Execute " Insert Into NewCaptions Values('33003','Yearly updation transfer leave balances for next year','33003') "
VstarConn.Execute " Insert Into NewCaptions Values('33004','Use only once at the beginning of the year','33004') "
VstarConn.Execute " Insert Into NewCaptions Values('33005','Leave Code','33005') "
VstarConn.Execute " Insert Into NewCaptions Values('33006','Leave Name','33006') "
VstarConn.Execute " Insert Into NewCaptions Values('33007','Please Wait...Updating Yearly Leaves','33007') "
VstarConn.Execute " Insert Into NewCaptions Values('33008','&' || 'Update','33008') "
VstarConn.Execute " Insert Into NewCaptions Values('33009','Yearly Leaves Already Updated ,Still Continue ?','33009') "
VstarConn.Execute " Insert Into NewCaptions Values('35001','Insert Memo Text','35001') "
VstarConn.Execute " Insert Into NewCaptions Values('35002','Memo','35002') "
VstarConn.Execute " Insert Into NewCaptions Values('35003','Ignore for','35003') "
VstarConn.Execute " Insert Into NewCaptions Values('36001','Monthly Process','36001') "
VstarConn.Execute " Insert Into NewCaptions Values('36002','Process for','36002') "
VstarConn.Execute " Insert Into NewCaptions Values('36003','All Employees','36003') "
VstarConn.Execute " Insert Into NewCaptions Values('36004','Selected Employees','36004') "
VstarConn.Execute " Insert Into NewCaptions Values('36005','Employees','36005') "
VstarConn.Execute " Insert Into NewCaptions Values('36006','Execute Late / Early rules','36006') "
VstarConn.Execute " Insert Into NewCaptions Values('36007','Consider Data from last month''s file','36007') "
VstarConn.Execute " Insert Into NewCaptions Values('36008','from the day','36008') "
VstarConn.Execute " Insert Into NewCaptions Values('36009','Onwards','36009') "
VstarConn.Execute " Insert Into NewCaptions Values('36010','&' || 'Process','36010') "
VstarConn.Execute " Insert Into NewCaptions Values('36011','Click this only for final Process.','36011') "
VstarConn.Execute " Insert Into NewCaptions Values('36012','If Processed again you may get false results.','36012') "
VstarConn.Execute " Insert Into NewCaptions Values('36013','Monthly Process Complete','36013') "
VstarConn.Execute " Insert Into NewCaptions Values('36014','     Please mention the day from which','36014') "
VstarConn.Execute " Insert Into NewCaptions Values('36015','the data from last month''s file has to be taken.','36015') "
VstarConn.Execute " Insert Into NewCaptions Values('36016',' Day cannot be greater than 31','36016') "
VstarConn.Execute " Insert Into NewCaptions Values('36017','Leave balance file for the year','36017') "
VstarConn.Execute " Insert Into NewCaptions Values('36018','Leave Information file for the year','36018') "
VstarConn.Execute " Insert Into NewCaptions Values('36019','Transaction file for the month','36019') "
VstarConn.Execute " Insert Into NewCaptions Values('36020','Difference between From Date and To Date cannot be more than 31','36020') "
VstarConn.Execute " Insert Into NewCaptions Values('36021','Difference between From Month and To Month cannot be more than 1','36021') "
VstarConn.Execute " Insert Into NewCaptions Values('36022',' does not exist in Leave Transaction File','36022') "
VstarConn.Execute " Insert Into NewCaptions Values('36023','Can not continue process.','36023') "
VstarConn.Execute " Insert Into NewCaptions Values('37001','Open Balance Entry','37001') "
VstarConn.Execute " Insert Into NewCaptions Values('37002','Opening on.','37002') "
VstarConn.Execute " Insert Into NewCaptions Values('37003','Please Select the Leave to be Added as Opening Balance','37003') "
VstarConn.Execute " Insert Into NewCaptions Values('37004','Opening Balance Leave(s) cannot be Added for 0 Number of days','37004') "
VstarConn.Execute " Insert Into NewCaptions Values('37005','Opening Balance Leave Days Must be Divisible by 0.50','37005') "
VstarConn.Execute " Insert Into NewCaptions Values('37006','Leaves already Added as Opening Balance on the above Selected Date','37006') "
VstarConn.Execute " Insert Into NewCaptions Values('40001','Reports','40001') "
VstarConn.Execute " Insert Into NewCaptions Values('40002','&' || 'Daily','40002') "
VstarConn.Execute " Insert Into NewCaptions Values('40003','Daily &' || 'Reports','40003') "
VstarConn.Execute " Insert Into NewCaptions Values('40004','Report for the date of','40004') "
VstarConn.Execute " Insert Into NewCaptions Values('40005','Shift Code','40005') "
VstarConn.Execute " Insert Into NewCaptions Values('40006','Physical Arrival','40006') "
VstarConn.Execute " Insert Into NewCaptions Values('40007','Absent','40007') "
VstarConn.Execute " Insert Into NewCaptions Values('40008','Continuous Absent','40008') "
VstarConn.Execute " Insert Into NewCaptions Values('40009','Late Arrival','40009') "
VstarConn.Execute " Insert Into NewCaptions Values('40010','Early Departure','40010') "
VstarConn.Execute " Insert Into NewCaptions Values('40011','Performance','40011') "
VstarConn.Execute " Insert Into NewCaptions Values('40012','Irregular','40012') "
VstarConn.Execute " Insert Into NewCaptions Values('40013','Entries','40013') "
VstarConn.Execute " Insert Into NewCaptions Values('40014','Shift Arrangement','40014') "
VstarConn.Execute " Insert Into NewCaptions Values('40015','Manpower','40015') "
VstarConn.Execute " Insert Into NewCaptions Values('40016','Out door duty','40016') "
VstarConn.Execute " Insert Into NewCaptions Values('40017','Summary','40017') "
VstarConn.Execute " Insert Into NewCaptions Values('40018','&' || 'Weekly','40018') "
VstarConn.Execute " Insert Into NewCaptions Values('40019','Weekly &' || 'Reports','40019') "
VstarConn.Execute " Insert Into NewCaptions Values('40020','Report for Week Starting From','40020') "
VstarConn.Execute " Insert Into NewCaptions Values('40021','Attendance','40021') "
VstarConn.Execute " Insert Into NewCaptions Values('40022','Shift Schedule','40022') "
VstarConn.Execute " Insert Into NewCaptions Values('40023','&' || 'Monthly','40023') "
VstarConn.Execute " Insert Into NewCaptions Values('40024','Monthly &' || 'Reports','40024') "
VstarConn.Execute " Insert Into NewCaptions Values('40025','Report for the month of','40025') "
VstarConn.Execute " Insert Into NewCaptions Values('40026','Muster Report','40026') "
VstarConn.Execute " Insert Into NewCaptions Values('40027','Monthly Present','40027') "
VstarConn.Execute " Insert Into NewCaptions Values('40028','Monthly Absent','40028') "
VstarConn.Execute " Insert Into NewCaptions Values('40029','Overtime Paid','40029') "
VstarConn.Execute " Insert Into NewCaptions Values('40030','Absent Memo','40030') "
VstarConn.Execute " Insert Into NewCaptions Values('40031','Late/Early/Absent','40031') "
VstarConn.Execute " Insert Into NewCaptions Values('40032','Leave Balance','40032') "
VstarConn.Execute " Insert Into NewCaptions Values('40033','Late Arrival Memo','40033') "
VstarConn.Execute " Insert Into NewCaptions Values('40034','Early Departure Memo','40034') "
VstarConn.Execute " Insert Into NewCaptions Values('40035','Leave Consumption','40035') "
VstarConn.Execute " Insert Into NewCaptions Values('40036','Total Lates','40036') "
VstarConn.Execute " Insert Into NewCaptions Values('40037','Total Earlys','40037') "
VstarConn.Execute " Insert Into NewCaptions Values('40038','WO on Holiday','40038') "
VstarConn.Execute " Insert Into NewCaptions Values('40039','&' || 'Yearly','40039') "
VstarConn.Execute " Insert Into NewCaptions Values('40040','Yearly &' || 'Reports','40040') "
VstarConn.Execute " Insert Into NewCaptions Values('40041','Report for the year of','40041') "
VstarConn.Execute " Insert Into NewCaptions Values('40042','Man Days','40042') "
VstarConn.Execute " Insert Into NewCaptions Values('40043','Present','40043') "
VstarConn.Execute " Insert Into NewCaptions Values('40044','Leave Information','40044') "
VstarConn.Execute " Insert Into NewCaptions Values('40045','M&' || 'asters','40045') "
VstarConn.Execute " Insert Into NewCaptions Values('40046','Master &' || 'Reports','40046') "
VstarConn.Execute " Insert Into NewCaptions Values('40047','Employee List','40047') "
VstarConn.Execute " Insert Into NewCaptions Values('40048','Employee Details','40048') "
VstarConn.Execute " Insert Into NewCaptions Values('40049','Left Employees','40049') "
VstarConn.Execute " Insert Into NewCaptions Values('40050','Rotational Shift','40050') "
VstarConn.Execute " Insert Into NewCaptions Values('40051','Holiday','40051') "
VstarConn.Execute " Insert Into NewCaptions Values('40052','P&' || 'eriodic','40052') "
VstarConn.Execute " Insert Into NewCaptions Values('40053','Periodic &' || 'Reports','40053') "
VstarConn.Execute " Insert Into NewCaptions Values('40054','Report for the period from','40054') "
VstarConn.Execute " Insert Into NewCaptions Values('40055','Reports Available for 30/31 Days only','40055') "
VstarConn.Execute " Insert Into NewCaptions Values('40056','Selectio&' || 'n','40056') "
VstarConn.Execute " Insert Into NewCaptions Values('40057','Employee','40057') "
VstarConn.Execute " Insert Into NewCaptions Values('40058','Group by','40058') "
VstarConn.Execute " Insert Into NewCaptions Values('40059','Department/Category','40059') "
VstarConn.Execute " Insert Into NewCaptions Values('40060','Start new page when group changes','40060') "
VstarConn.Execute " Insert Into NewCaptions Values('40061','Print Date &' || '&' || ' Time','40061') "
VstarConn.Execute " Insert Into NewCaptions Values('40062','Use 132 Column Dot matrix Printer','40062') "
VstarConn.Execute " Insert Into NewCaptions Values('40063','Prompt before Printing','40063') "
VstarConn.Execute " Insert Into NewCaptions Values('40064','Please confirm that your default Printer is set to 132 column Dot Matrix Printer.','40064') "
VstarConn.Execute " Insert Into NewCaptions Values('40065','Monthly Transactin File not found for the Month of','40065') "
VstarConn.Execute " Insert Into NewCaptions Values('40066','Cannot Send Reports through Email :: Refer Install->Parameters','40066') "
VstarConn.Execute " Insert Into NewCaptions Values('40067','Perform printing?','40067') "
VstarConn.Execute " Insert Into NewCaptions Values('40068','Please Enter the Date First','40068') "
VstarConn.Execute " Insert Into NewCaptions Values('40069','Daily Process Required. Continue ?','40069') "
VstarConn.Execute " Insert Into NewCaptions Values('40070','If Monthly process for selected month is not done,','40070') "
VstarConn.Execute " Insert Into NewCaptions Values('40071','please do it first.','40071') "
VstarConn.Execute " Insert Into NewCaptions Values('40072','Network Problem : Please Retry','40072') "
VstarConn.Execute " Insert Into NewCaptions Values('40073','The Papersize of the selected Printer is smaller than','40073') "
VstarConn.Execute " Insert Into NewCaptions Values('40074','      the required Papersize for the Report.','40074') "
VstarConn.Execute " Insert Into NewCaptions Values('40075','Operation is cancelled .','40075') "
VstarConn.Execute " Insert Into NewCaptions Values('40076','Shift File not found for the Month of','40076') "
VstarConn.Execute " Insert Into NewCaptions Values('40077','Month can not be empty','40077') "
VstarConn.Execute " Insert Into NewCaptions Values('40078','Year can not be empty','40078') "
VstarConn.Execute " Insert Into NewCaptions Values('40079','Period should not be more than 31','40079') "
VstarConn.Execute " Insert Into NewCaptions Values('40080','Period should not be for more than 2 Months','40080') "
VstarConn.Execute " Insert Into NewCaptions Values('40081','Yearly Leave Transaction File Not Found','40081') "
VstarConn.Execute " Insert Into NewCaptions Values('40082','Yearly Leave Information File Not Found','40082') "
VstarConn.Execute " Insert Into NewCaptions Values('40083','Mail Address not found for','40083') "
VstarConn.Execute " Insert Into NewCaptions Values('40084','Meal Allowance','40084') "
VstarConn.Execute " Insert Into NewCaptions Values('40085','Punches Report','40085') "
VstarConn.Execute " Insert Into NewCaptions Values('40086','Permission cards','40086') "
VstarConn.Execute " Insert Into NewCaptions Values('41001','Rotation Master','41001') "
VstarConn.Execute " Insert Into NewCaptions Values('41002','Rotation Code','41002') "
VstarConn.Execute " Insert Into NewCaptions Values('41003','Shift Rotates','41003') "
VstarConn.Execute " Insert Into NewCaptions Values('41004','Only after specified number of days','41004') "
VstarConn.Execute " Insert Into NewCaptions Values('41005','On the following dates of every month','41005') "
VstarConn.Execute " Insert Into NewCaptions Values('41006','On following week days(SUN..SAT)','41006') "
VstarConn.Execute " Insert Into NewCaptions Values('41007','The shift changes from one to another','41007') "
VstarConn.Execute " Insert Into NewCaptions Values('41008','Rotation','41008') "
VstarConn.Execute " Insert Into NewCaptions Values('41009','Shift Rotation','41009') "
VstarConn.Execute " Insert Into NewCaptions Values('41010','Rotation Days/Dates','41010') "
VstarConn.Execute " Insert Into NewCaptions Values('41011','Rotation Code Already Exists','41011') "
VstarConn.Execute " Insert Into NewCaptions Values('41012','Rotation Name cannot be Blank','41012') "
VstarConn.Execute " Insert Into NewCaptions Values('41013','Days are not Specified , Please Do It','41013') "
VstarConn.Execute " Insert Into NewCaptions Values('41014','Dates are not Specified , Please Do It','41014') "
VstarConn.Execute " Insert Into NewCaptions Values('41015','Week Days are not Specified , Please Do It','41015') "
VstarConn.Execute " Insert Into NewCaptions Values('41016','Shifts are not Specified , Please Do It','41016') "
VstarConn.Execute " Insert Into NewCaptions Values('41017','Rotation with Same Code Already Exists','41017') "
VstarConn.Execute " Insert Into NewCaptions Values('41018','Click to Select the Number of Days','41018') "
VstarConn.Execute " Insert Into NewCaptions Values('41019','Click to Select the Dates of Every Month','41019') "

VstarConn.Execute " Insert Into NewCaptions Values('41020','Click to Select the Week','41020') "
VstarConn.Execute " Insert Into NewCaptions Values('41021','Click to Select the Shiffts','41021') "
VstarConn.Execute " Insert Into NewCaptions Values('41022','Rotation Code cannot be blank','41022') "
VstarConn.Execute " Insert Into NewCaptions Values('42001','Only after specified number of days','42001') "
VstarConn.Execute " Insert Into NewCaptions Values('43001','Select Week Days','43001') "
VstarConn.Execute " Insert Into NewCaptions Values('44001','Late /Early Rules','44001') "
VstarConn.Execute " Insert Into NewCaptions Values('44002','Select C&' || 'ategory','44002') "
VstarConn.Execute " Insert Into NewCaptions Values('44003','Late Deductions','44003') "
VstarConn.Execute " Insert Into NewCaptions Values('44004','Total Late Allowed in a Month','44004') "
VstarConn.Execute " Insert Into NewCaptions Values('44005','Cut','44005') "
VstarConn.Execute " Insert Into NewCaptions Values('44006','Day for every','44006') "
VstarConn.Execute " Insert Into NewCaptions Values('44007','&' || 'Deduct From','44007') "
VstarConn.Execute " Insert Into NewCaptions Values('44008','Paid Days','44008') "
VstarConn.Execute " Insert Into NewCaptions Values('44009','Leaves','44009') "
VstarConn.Execute " Insert Into NewCaptions Values('44010','1st Preference','44010') "
VstarConn.Execute " Insert Into NewCaptions Values('44011','2nd Preference','44011') "
VstarConn.Execute " Insert Into NewCaptions Values('44012','3rd Preference','44012') "
VstarConn.Execute " Insert Into NewCaptions Values('44013','Early Deductions','44013') "
VstarConn.Execute " Insert Into NewCaptions Values('44014','Total Early Allowed in Month','44014') "
VstarConn.Execute " Insert Into NewCaptions Values('44015','&' || 'Reset','44015') "
VstarConn.Execute " Insert Into NewCaptions Values('44016','Are You Sure to Reset the Rules','44016') "
VstarConn.Execute " Insert Into NewCaptions Values('44017','Days To be Cut Must be Divisible by 0.50','44017') "
VstarConn.Execute " Insert Into NewCaptions Values('44018','Fractional Number not Allowed','44018') "
VstarConn.Execute " Insert Into NewCaptions Values('44019','Please Select atleast 1 Leave','44019') "
VstarConn.Execute " Insert Into NewCaptions Values('44020','Please Select Leave for the Second Preference','44020') "
VstarConn.Execute " Insert Into NewCaptions Values('44021','Please Select Leave for the First Preference','44021') "
VstarConn.Execute " Insert Into NewCaptions Values('45001','Shift schedule for','45001') "
VstarConn.Execute " Insert Into NewCaptions Values('45002','1st Week','45002') "
VstarConn.Execute " Insert Into NewCaptions Values('45003','2nd Week','45003') "
VstarConn.Execute " Insert Into NewCaptions Values('45004','3rd Week','45004') "
VstarConn.Execute " Insert Into NewCaptions Values('45005','4th Week','45005') "
VstarConn.Execute " Insert Into NewCaptions Values('45006','5th Week','45006') "
VstarConn.Execute " Insert Into NewCaptions Values('45007','&' || 'Period','45007') "
VstarConn.Execute " Insert Into NewCaptions Values('45008','M&' || 'aster','45008') "
VstarConn.Execute " Insert Into NewCaptions Values('45009','MON','45009') "
VstarConn.Execute " Insert Into NewCaptions Values('45010','TUE','45010') "
VstarConn.Execute " Insert Into NewCaptions Values('45011','WED','45011') "
VstarConn.Execute " Insert Into NewCaptions Values('45012','THU','45012') "
VstarConn.Execute " Insert Into NewCaptions Values('45013','FRI','45013') "
VstarConn.Execute " Insert Into NewCaptions Values('45014','SAT','45014') "
VstarConn.Execute " Insert Into NewCaptions Values('45015','SUN','45015') "
VstarConn.Execute " Insert Into NewCaptions Values('45016','Monthly Shift File not Found for the Month of','45016') "
VstarConn.Execute " Insert Into NewCaptions Values('45017','Shift for the Employee','45017') "
VstarConn.Execute " Insert Into NewCaptions Values('45018',' not Yet Created :: Please Create it','45018') "
VstarConn.Execute " Insert Into NewCaptions Values('46001','Select Dat File','46001') "
VstarConn.Execute " Insert Into NewCaptions Values('46002','&' || 'Files','46002') "
VstarConn.Execute " Insert Into NewCaptions Values('46003','Fold&' || 'ers','46003') "
VstarConn.Execute " Insert Into NewCaptions Values('46004','&' || 'Drives','46004') "
VstarConn.Execute " Insert Into NewCaptions Values('47001','Select Shifts','47001') "
VstarConn.Execute " Insert Into NewCaptions Values('47002','Start','47002') "
VstarConn.Execute " Insert Into NewCaptions Values('47003','End','47003') "
VstarConn.Execute " Insert Into NewCaptions Values('47004','Night','47004') "
VstarConn.Execute " Insert Into NewCaptions Values('47005','&' || 'Reset','47005') "
VstarConn.Execute " Insert Into NewCaptions Values('48001','Shift Master','48001') "
VstarConn.Execute " Insert Into NewCaptions Values('48002','Shift Code','48002') "
VstarConn.Execute " Insert Into NewCaptions Values('48003','This is a night shift','48003') "
VstarConn.Execute " Insert Into NewCaptions Values('48004','Deduct Break Hrs from Shift hrs','48004') "
VstarConn.Execute " Insert Into NewCaptions Values('48005','Shift Time','48005') "
VstarConn.Execute " Insert Into NewCaptions Values('48006','Shift starts at','48006') "
VstarConn.Execute " Insert Into NewCaptions Values('48007','First half ends at','48007') "
VstarConn.Execute " Insert Into NewCaptions Values('48008','Second half starts at','48008') "
VstarConn.Execute " Insert Into NewCaptions Values('48009','Shift Ends at','48009') "
VstarConn.Execute " Insert Into NewCaptions Values('48010','Total shift time','48010') "
VstarConn.Execute " Insert Into NewCaptions Values('48011','Break Periods','48011') "
VstarConn.Execute " Insert Into NewCaptions Values('48012','Starts at','48012') "
VstarConn.Execute " Insert Into NewCaptions Values('48013','Ends at','48013') "
VstarConn.Execute " Insert Into NewCaptions Values('48014','Break','48014') "
VstarConn.Execute " Insert Into NewCaptions Values('48015','First Break','48015') "
VstarConn.Execute " Insert Into NewCaptions Values('48016','Second Break','48016') "
VstarConn.Execute " Insert Into NewCaptions Values('48017','Third Break','48017') "
VstarConn.Execute " Insert Into NewCaptions Values('48018','Shift Code cannot be blank','48018') "
VstarConn.Execute " Insert Into NewCaptions Values('48019','Shift Already Exists','48019') "
VstarConn.Execute " Insert Into NewCaptions Values('48020','Shift Name cannot be blank','48020') "
VstarConn.Execute " Insert Into NewCaptions Values('48021','Second Shift Start time cannot Be Less than First Shift End Time','48021') "
VstarConn.Execute " Insert Into NewCaptions Values('48022','First Break End Time cannot be Less than First Break Start Time','48022') "
VstarConn.Execute " Insert Into NewCaptions Values('48023','Second Break End Time cannot be Less than Second Break Start Time','48023') "
VstarConn.Execute " Insert Into NewCaptions Values('48024','Third Break End Time cannot be Less than Third Break Start Time','48024') "
VstarConn.Execute " Insert Into NewCaptions Values('48025','Second Break Start Time Cannot be Less than First Break End Time','48025') "
VstarConn.Execute " Insert Into NewCaptions Values('48026','Third Break Start Time Cannot be Less than Second Break End Time','48026') "
VstarConn.Execute " Insert Into NewCaptions Values('48027','Time Should be Greater than Shift Start Time','48027') "
VstarConn.Execute " Insert Into NewCaptions Values('48028','Time Should be Less than Shift End Time','48028') "
VstarConn.Execute " Insert Into NewCaptions Values('48029','Shift with Same Code Already Exists','48029') "
VstarConn.Execute " Insert Into NewCaptions Values('48030','Shift not Found','48030') "
VstarConn.Execute " Insert Into NewCaptions Values('48031','Are you Sure to Delete this Record','48031') "
VstarConn.Execute " Insert Into NewCaptions Values('49001','Monthly shift creation','49001') "
VstarConn.Execute " Insert Into NewCaptions Values('49004','Shift Date','49004') "
VstarConn.Execute " Insert Into NewCaptions Values('49005',':: Please Wait...Processing Shifts','49005') "
VstarConn.Execute " Insert Into NewCaptions Values('49006','Shift file processed successfully for the month of','49006') "
VstarConn.Execute " Insert Into NewCaptions Values('49007','Shift file Already Exists for the Month of','49007') "
VstarConn.Execute " Insert Into NewCaptions Values('49008',' Do You Wish to Overwrite it','49008') "
VstarConn.Execute " Insert Into NewCaptions Values('49009','Employee Selection','49009') "
VstarConn.Execute " Insert Into NewCaptions Values('49010','Click to Toggle Selection','49010') "
VstarConn.Execute " Insert Into NewCaptions Values('49011','Create monthly shift schedule','49011') "
VstarConn.Execute " Insert Into NewCaptions Values('50001','Select Shift','50001') "
VstarConn.Execute " Insert Into NewCaptions Values('50002','Start','50002') "
VstarConn.Execute " Insert Into NewCaptions Values('50003','End','50003') "
VstarConn.Execute " Insert Into NewCaptions Values('50004','Night','50004') "
VstarConn.Execute " Insert Into NewCaptions Values('50005','Please Select the Shift','50005') "
VstarConn.Execute " Insert Into NewCaptions Values('52001','Login','52001') "
VstarConn.Execute " Insert Into NewCaptions Values('52002','Password','52002') "
VstarConn.Execute " Insert Into NewCaptions Values('52003','User Name','52003') "
VstarConn.Execute " Insert Into NewCaptions Values('52004','Exclusive Access','52004') "
VstarConn.Execute " Insert Into NewCaptions Values('52005','User Name cannot be Blank','52005') "
VstarConn.Execute " Insert Into NewCaptions Values('52006','There is an Exclusive Lock by the Administrator :: Cannot Start the Application','52006') "
VstarConn.Execute " Insert Into NewCaptions Values('52007','Invalid Username or Password :: Try Again','52007') "
VstarConn.Execute " Insert Into NewCaptions Values('53001','&' || 'Install','53001') "
VstarConn.Execute " Insert Into NewCaptions Values('53002','Pa&' || 'rameter','53002') "
VstarConn.Execute " Insert Into NewCaptions Values('53003','&' || 'Shift','53003') "
VstarConn.Execute " Insert Into NewCaptions Values('53004','C&' || 'ompany','53004') "
VstarConn.Execute " Insert Into NewCaptions Values('53005','&' || 'Category','53005') "
VstarConn.Execute " Insert Into NewCaptions Values('53006','&' || 'Leave','53006') "
VstarConn.Execute " Insert Into NewCaptions Values('53007','&' || 'Yearly Leaves','53007') "
VstarConn.Execute " Insert Into NewCaptions Values('53008','&' || 'Update','53008') "
VstarConn.Execute " Insert Into NewCaptions Values('53009','&' || 'Rules','53009') "
VstarConn.Execute " Insert Into NewCaptions Values('53010','Login as &' || 'Different User','53010') "
VstarConn.Execute " Insert Into NewCaptions Values('53011','&' || 'Updation','53011') "
VstarConn.Execute " Insert Into NewCaptions Values('53012','&' || 'Department Master','53012') "
VstarConn.Execute " Insert Into NewCaptions Values('53013','&' || 'Group Master','53013') "
VstarConn.Execute " Insert Into NewCaptions Values('53014','&' || 'Rotation Shift Master','53014') "
VstarConn.Execute " Insert Into NewCaptions Values('53015','&' || 'Holidays','53015') "
VstarConn.Execute " Insert Into NewCaptions Values('53016','Ho&' || 'liday Master','53016') "
VstarConn.Execute " Insert Into NewCaptions Values('53017','&' || 'Declare Holiday','53017') "
VstarConn.Execute " Insert Into NewCaptions Values('53018','&' || 'Employee Master','53018') "
VstarConn.Execute " Insert Into NewCaptions Values('53019','&' || 'Shift Schedule','53019') "
VstarConn.Execute " Insert Into NewCaptions Values('53020','&' || 'Schedule Master','53020') "
VstarConn.Execute " Insert Into NewCaptions Values('53021','&' || 'Create Schedule','53021') "
VstarConn.Execute " Insert Into NewCaptions Values('53022','Chan&' || 'ge Schedule','53022') "
VstarConn.Execute " Insert Into NewCaptions Values('53023','&' || 'Leaves Transaction','53023') "
VstarConn.Execute " Insert Into NewCaptions Values('53024','&' || 'Opening','53024') "
VstarConn.Execute " Insert Into NewCaptions Values('53025','&' || 'Credit','53025') "
VstarConn.Execute " Insert Into NewCaptions Values('53026','&' || 'Encash','53026') "
VstarConn.Execute " Insert Into NewCaptions Values('53027','&' || 'Avail','53027') "
VstarConn.Execute " Insert Into NewCaptions Values('53028','Lo&' || 'st Entry','53028') "
VstarConn.Execute " Insert Into NewCaptions Values('53029','C&' || 'orrection','53029') "
VstarConn.Execute " Insert Into NewCaptions Values('53030','Edit &' || 'Paid Days','53030') "
VstarConn.Execute " Insert Into NewCaptions Values('53031','&' || 'Process','53031') "
VstarConn.Execute " Insert Into NewCaptions Values('53032','&' || 'Daily','53032') "
VstarConn.Execute " Insert Into NewCaptions Values('53033','&' || 'Monthly','53033') "
VstarConn.Execute " Insert Into NewCaptions Values('53034','&' || 'Report','53034') "
VstarConn.Execute " Insert Into NewCaptions Values('53035','&' || 'Reports','53035') "
VstarConn.Execute " Insert Into NewCaptions Values('53036','U&' || 'tility','53036') "
VstarConn.Execute " Insert Into NewCaptions Values('53037','Change &' || 'Password','53037') "
VstarConn.Execute " Insert Into NewCaptions Values('53038','&' || 'Backup and Restore','53038') "
VstarConn.Execute " Insert Into NewCaptions Values('53039','&' || 'Reset Locks','53039') "
VstarConn.Execute " Insert Into NewCaptions Values('53040','Compact &' || 'Database','53040') "
VstarConn.Execute " Insert Into NewCaptions Values('53041','&' || 'Version','53041') "
VstarConn.Execute " Insert Into NewCaptions Values('53042','&' || 'About','53042') "
VstarConn.Execute " Insert Into NewCaptions Values('53043','Support','53043') "
VstarConn.Execute " Insert Into NewCaptions Values('53044','User &' || 'Accounts','53044') "
VstarConn.Execute " Insert Into NewCaptions Values('53045','Monthly Data is  in Process : Please Try after Some Time','53045') "
VstarConn.Execute " Insert Into NewCaptions Values('53046','Yearly Leave Files are in Process :: Please Try after some Time','53046') "
VstarConn.Execute " Insert Into NewCaptions Values('53047','Data is  in Process : Please Try after Some Time','53047') "
VstarConn.Execute " Insert Into NewCaptions Values('53048','&' || 'Create','53048') "
VstarConn.Execute " Insert Into NewCaptions Values('53049','&' || 'Language','53049') "
VstarConn.Execute " Insert Into NewCaptions Values('53050','&' || 'Switch Language','53050') "
VstarConn.Execute " Insert Into NewCaptions Values('53051','&' || 'Translate','53051') "
VstarConn.Execute " Insert Into NewCaptions Values('53052','&' || 'Report Captions','53052') "
VstarConn.Execute " Insert Into NewCaptions Values('53053','&' || 'General Captions','53053') "
VstarConn.Execute " Insert Into NewCaptions Values('53054','&' || 'Change Menu Captions','53054') "
VstarConn.Execute " Insert Into NewCaptions Values('53055','E&' || 'xit','53055') "
VstarConn.Execute " Insert Into NewCaptions Values('53056','&' || 'Daily Process','53056') "
VstarConn.Execute " Insert Into NewCaptions Values('53057','&' || 'Monthly Process','53057') "
VstarConn.Execute " Insert Into NewCaptions Values('53058','O&' || 'T Rules','53058') "
VstarConn.Execute " Insert Into NewCaptions Values('53059','C&' || 'O Rules','53059') "
VstarConn.Execute " Insert Into NewCaptions Values('53060','C&' || 'hange User Password','53060') "
VstarConn.Execute " Insert Into NewCaptions Values('53061','OT &' || 'Authorization','53061') "
VstarConn.Execute " Insert Into NewCaptions Values('53062','Loc&' || 'ation Master','53062') "
VstarConn.Execute " Insert Into NewCaptions Values('53063','&' || 'Set Employee Details','53063') "
VstarConn.Execute " Insert Into NewCaptions Values('53064','&' || 'Division Master','53064') "
VstarConn.Execute " Insert Into NewCaptions Values('53065','&' || 'Export Data','53065') "
'' apoorva
VstarConn.Execute " Insert Into NewCaptions Values('53066','&Dos Reports','53066') "
VstarConn.Execute " Insert Into NewCaptions Values('54001','Installation Parameters','54001') "
VstarConn.Execute " Insert Into NewCaptions Values('54002','Parameters','54002') "
VstarConn.Execute " Insert Into NewCaptions Values('54003','General','54003') "
VstarConn.Execute " Insert Into NewCaptions Values('54004','Current year is','54004') "
VstarConn.Execute " Insert Into NewCaptions Values('54005','Employee code size is','54005') "
VstarConn.Execute " Insert Into NewCaptions Values('54006','Punching card size is','54006') "
VstarConn.Execute " Insert Into NewCaptions Values('54007','Year Starting from','54007') "
VstarConn.Execute " Insert Into NewCaptions Values('54008','Week begins on','54008') "
VstarConn.Execute " Insert Into NewCaptions Values('54009','Employee can work after the shift for next','54009') "
VstarConn.Execute " Insert Into NewCaptions Values('54010','Select Path for .DAT File','54010') "
VstarConn.Execute " Insert Into NewCaptions Values('54011','Ignore next punch from  the previous punch till','54011') "
VstarConn.Execute " Insert Into NewCaptions Values('54012','Allow Employee to be posted to next shift if late by','54012') "
VstarConn.Execute " Insert Into NewCaptions Values('54013','Allow Employee to be posted to Previous shift if Early by','54013') "
VstarConn.Execute " Insert Into NewCaptions Values('54014','Details','54014') "
VstarConn.Execute " Insert Into NewCaptions Values('54015','Permission Cards','54015') "
VstarConn.Execute " Insert Into NewCaptions Values('54016','Use Permission cards','54016') "
VstarConn.Execute " Insert Into NewCaptions Values('54017','Starting number','54017') "
VstarConn.Execute " Insert Into NewCaptions Values('54018','Ending number','54018') "
VstarConn.Execute " Insert Into NewCaptions Values('54019','Card number','54019') "
VstarConn.Execute " Insert Into NewCaptions Values('54020','Late Coming','54020') "
VstarConn.Execute " Insert Into NewCaptions Values('54021','Early Going','54021') "
VstarConn.Execute " Insert Into NewCaptions Values('54022','Late bus','54022') "
VstarConn.Execute " Insert Into NewCaptions Values('54023','Official Duty','54023') "
VstarConn.Execute " Insert Into NewCaptions Values('54024','OverTime','54024') "
VstarConn.Execute " Insert Into NewCaptions Values('54025','Rates','54025') "
VstarConn.Execute " Insert Into NewCaptions Values('54026','On a Holiday, Calculate @','54026') "
VstarConn.Execute " Insert Into NewCaptions Values('54027','On a Week Off,Calculate @','54027') "
VstarConn.Execute " Insert Into NewCaptions Values('54028','On a Working Day,Calculate @','54028') "
VstarConn.Execute " Insert Into NewCaptions Values('54029','times, the  total work hours for that day','54029') "
VstarConn.Execute " Insert Into NewCaptions Values('54030','Deduct Late / Early Hours from OverTime','54030') "

VstarConn.Execute " Insert Into NewCaptions Values('54031','Roundoff Over Time Hours','54031') "
VstarConn.Execute " Insert Into NewCaptions Values('54032','&' || 'Back','54032') "
VstarConn.Execute " Insert Into NewCaptions Values('54033','&' || 'Next','54033') "
VstarConn.Execute " Insert Into NewCaptions Values('54034','Select Path for .DAT File','54034') "
VstarConn.Execute " Insert Into NewCaptions Values('54035','digit','54035') "
VstarConn.Execute " Insert Into NewCaptions Values('54036','Send Reports using Email','54036') "
VstarConn.Execute " Insert Into NewCaptions Values('54037','Overtime Rounding','54037') "
VstarConn.Execute " Insert Into NewCaptions Values('54038','Decimal Value (in Minutes)','54038') "
VstarConn.Execute " Insert Into NewCaptions Values('54039','OT Round off Pattern','54039') "
VstarConn.Execute " Insert Into NewCaptions Values('54040','Round off to','54040') "
VstarConn.Execute " Insert Into NewCaptions Values('54041','More than','54041') "
VstarConn.Execute " Insert Into NewCaptions Values('54042','Apply Salary Cut-Off Date','54042') "
VstarConn.Execute " Insert Into NewCaptions Values('54043','Cut-Off Date','54043') "
VstarConn.Execute " Insert Into NewCaptions Values('54044','Application Date Format :','54044') "
VstarConn.Execute " Insert Into NewCaptions Values('54045','Size Once Increased cannot be Decreased','54045') "
VstarConn.Execute " Insert Into NewCaptions Values('54046','Warning : Size should not be decreased','54046') "
VstarConn.Execute " Insert Into NewCaptions Values('54047','This value should be greater than','54047') "
VstarConn.Execute " Insert Into NewCaptions Values('54048','Cut-Off Date Cannot be Greater than 31','54048') "
VstarConn.Execute " Insert Into NewCaptions Values('54049','Please enter appropriate Values in the OD Card','54049') "
VstarConn.Execute " Insert Into NewCaptions Values('54050','Please enter appropriate Values in the Late Card','54050') "
VstarConn.Execute " Insert Into NewCaptions Values('54051','Please enter appropriate Values in the Late Bus Card','54051') "
VstarConn.Execute " Insert Into NewCaptions Values('54052','Please enter appropriate Values in the Early  Card','54052') "
VstarConn.Execute " Insert Into NewCaptions Values('54053','Please enter appropriate Values in the Start Card','54053') "
VstarConn.Execute " Insert Into NewCaptions Values('54054','Please enter appropriate Values in the End Card','54054') "
VstarConn.Execute " Insert Into NewCaptions Values('54055','OFF DUTY card no should be in the range of','54055') "
VstarConn.Execute " Insert Into NewCaptions Values('54056','LATE card no should be in the range of','54056') "
VstarConn.Execute " Insert Into NewCaptions Values('54057','EARLY card no should be in the range of','54057') "
VstarConn.Execute " Insert Into NewCaptions Values('54058','LATE BUS card no should be in the range of','54058') "
VstarConn.Execute " Insert Into NewCaptions Values('54059','Two cards are given same number. Cannot save the parameters.','54059') "
VstarConn.Execute " Insert Into NewCaptions Values('54060','If Salary Generation is after Month Completion then keep CutOff Day as 0','54060') "
VstarConn.Execute " Insert Into NewCaptions Values('54061','Maximum value can be 10.00','54061') "
VstarConn.Execute " Insert Into NewCaptions Values('56001','User Accounts','56001') "
VstarConn.Execute " Insert Into NewCaptions Values('56002','List','56002') "
VstarConn.Execute " Insert Into NewCaptions Values('56003','Details','56003') "
VstarConn.Execute " Insert Into NewCaptions Values('56004','Password / UserName','56004') "
VstarConn.Execute " Insert Into NewCaptions Values('56005','Master File Rights','56005') "
VstarConn.Execute " Insert Into NewCaptions Values('56006','Other Rights','56006') "
VstarConn.Execute " Insert Into NewCaptions Values('56007','Leave Transaction','56007') "
VstarConn.Execute " Insert Into NewCaptions Values('56008','Select/Unselect All','56008') "
VstarConn.Execute " Insert Into NewCaptions Values('56009','Parameter','56009') "
VstarConn.Execute " Insert Into NewCaptions Values('56010','Process','56010') "
VstarConn.Execute " Insert Into NewCaptions Values('56011','Yearly Leaves','56011') "
VstarConn.Execute " Insert Into NewCaptions Values('56012','Edit','56012') "
VstarConn.Execute " Insert Into NewCaptions Values('56013','Daily','56013') "
VstarConn.Execute " Insert Into NewCaptions Values('56014','Monthly','56014') "
VstarConn.Execute " Insert Into NewCaptions Values('56015','Create','56015') "
VstarConn.Execute " Insert Into NewCaptions Values('56016','Update','56016') "
VstarConn.Execute " Insert Into NewCaptions Values('56017','Change','56017') "
VstarConn.Execute " Insert Into NewCaptions Values('56018','Record','56018') "
VstarConn.Execute " Insert Into NewCaptions Values('56019','On duty','56019') "
VstarConn.Execute " Insert Into NewCaptions Values('56020','Off Duty','56020') "
VstarConn.Execute " Insert Into NewCaptions Values('56021','OT','56021') "
VstarConn.Execute " Insert Into NewCaptions Values('56022','Time','56022') "
VstarConn.Execute " Insert Into NewCaptions Values('56023','Add/Edit/Delete of Login Users','56023') "
VstarConn.Execute " Insert Into NewCaptions Values('56024','Backup','56024') "
VstarConn.Execute " Insert Into NewCaptions Values('56025','Restore','56025') "
VstarConn.Execute " Insert Into NewCaptions Values('56026','Login users','56026') "
VstarConn.Execute " Insert Into NewCaptions Values('56027','Correction','56027') "
VstarConn.Execute " Insert Into NewCaptions Values('56028','Shift Schedule','56028') "
VstarConn.Execute " Insert Into NewCaptions Values('56029','Paid Days','56029') "
VstarConn.Execute " Insert Into NewCaptions Values('56030','User Code','56030') "
VstarConn.Execute " Insert Into NewCaptions Values('56031','Password','56031') "
VstarConn.Execute " Insert Into NewCaptions Values('56032','Credit','56032') "
VstarConn.Execute " Insert Into NewCaptions Values('56033','Encash','56033') "
VstarConn.Execute " Insert Into NewCaptions Values('56034','Avail','56034') "
VstarConn.Execute " Insert Into NewCaptions Values('56035','Add','56035') "
VstarConn.Execute " Insert Into NewCaptions Values('56036','Delete','56036') "
VstarConn.Execute " Insert Into NewCaptions Values('56037','Opening','56037') "
VstarConn.Execute " Insert Into NewCaptions Values('56038','Rules','56038') "
VstarConn.Execute " Insert Into NewCaptions Values('56039','Compact Database','56039') "
VstarConn.Execute " Insert Into NewCaptions Values('56040','Permission','56040') "
VstarConn.Execute " Insert Into NewCaptions Values('56041','Maximum users allowed :','56041') "
VstarConn.Execute " Insert Into NewCaptions Values('56042','User Name cannot be blank.','56042') "
VstarConn.Execute " Insert Into NewCaptions Values('56043','Password cannot be blank.','56043') "
VstarConn.Execute " Insert Into NewCaptions Values('56044','Encryption Error :: Try Another Password','56044') "
VstarConn.Execute " Insert Into NewCaptions Values('56045','Unable to change the Password','56045') "
VstarConn.Execute " Insert Into NewCaptions Values('56046','Error changing the Password','56046') "
VstarConn.Execute " Insert Into NewCaptions Values('56047','The User Cannot Delete Himself','56047') "
VstarConn.Execute " Insert Into NewCaptions Values('56048','Are You Sure To Delete User','56048') "
VstarConn.Execute " Insert Into NewCaptions Values('56049','Please Select the User','56049') "
VstarConn.Execute " Insert Into NewCaptions Values('56050','The User Already Exists','56050') "
VstarConn.Execute " Insert Into NewCaptions Values('56051','No Records Obtained From User Master','56051') "
VstarConn.Execute " Insert Into NewCaptions Values('56052','Re-Start the Application','56052') "
VstarConn.Execute " Insert Into NewCaptions Values('56053','You are about to change your login password','56053') "
VstarConn.Execute " Insert Into NewCaptions Values('56054','Do you want to proceed ?','56054') "
VstarConn.Execute " Insert Into NewCaptions Values('56055','Your Login Password has been changed successfully','56055') "
VstarConn.Execute " Insert Into NewCaptions Values('56056','Menu Items','56056') "
VstarConn.Execute " Insert Into NewCaptions Values('56057','User Name','56057') "
VstarConn.Execute " Insert Into NewCaptions Values('56058','BackUp and Restore','56058') "
VstarConn.Execute " Insert Into NewCaptions Values('56059','Press Spacebar or Double Click to Toggle Rights','56059') "
VstarConn.Execute " Insert Into NewCaptions Values('57001','Visual star Version','57001') "
VstarConn.Execute " Insert Into NewCaptions Values('57002','Version information','57002') "
VstarConn.Execute " Insert Into NewCaptions Values('57003','Comments','57003') "
VstarConn.Execute " Insert Into NewCaptions Values('57004','Company Name','57004') "
VstarConn.Execute " Insert Into NewCaptions Values('57005','File Description','57005') "
VstarConn.Execute " Insert Into NewCaptions Values('57006','File Version','57006') "
VstarConn.Execute " Insert Into NewCaptions Values('57007','Internal Name','57007') "
VstarConn.Execute " Insert Into NewCaptions Values('57008','Legal copyright','57008') "
VstarConn.Execute " Insert Into NewCaptions Values('57009','Legal trademarks','57009') "
VstarConn.Execute " Insert Into NewCaptions Values('57010','Original Filename','57010') "
VstarConn.Execute " Insert Into NewCaptions Values('57011','Product Name','57011') "
VstarConn.Execute " Insert Into NewCaptions Values('57012','Product Version','57012') "
VstarConn.Execute " Insert Into NewCaptions Values('57013','Special Build for','57013') "
VstarConn.Execute " Insert Into NewCaptions Values('58001','Yearly File Creation','58001') "
VstarConn.Execute " Insert Into NewCaptions Values('58002','Yearly file creation for the selected Year','58002') "
VstarConn.Execute " Insert Into NewCaptions Values('58003','Use this option :','58003') "
VstarConn.Execute " Insert Into NewCaptions Values('58004','=> In the beginning of the year','58004') "
VstarConn.Execute " Insert Into NewCaptions Values('58005','=> When new type of leave is Added','58005') "
VstarConn.Execute " Insert Into NewCaptions Values('58006','=> When existing leave is Edited or Deleted','58006') "
VstarConn.Execute " Insert Into NewCaptions Values('58007','Instructions','58007') "
VstarConn.Execute " Insert Into NewCaptions Values('58008','&' || 'Create','58008') "
VstarConn.Execute " Insert Into NewCaptions Values('58009','File :','58009') "
VstarConn.Execute " Insert Into NewCaptions Values('58010','Already exists, You want to overwrite ?','58010') "
VstarConn.Execute " Insert Into NewCaptions Values('58011','Yearly Leave Files Creation finished','58011') "
VstarConn.Execute " Insert Into NewCaptions Values('59001','OT Rule No.','59001') "
VstarConn.Execute " Insert Into NewCaptions Values('59002','Give OT On','59002') "
VstarConn.Execute " Insert Into NewCaptions Values('59003','More than','59003') "
VstarConn.Execute " Insert Into NewCaptions Values('59004','Apply deduction on Weekoff','59004') "
VstarConn.Execute " Insert Into NewCaptions Values('59005','Apply deduction on Holiday','59005') "
VstarConn.Execute " Insert Into NewCaptions Values('59006','','59006') "
VstarConn.Execute " Insert Into NewCaptions Values('59007','Authorized by default','59007') "
VstarConn.Execute " Insert Into NewCaptions Values('59008','Maximum OT can be upto','59008') "
VstarConn.Execute " Insert Into NewCaptions Values('59009','Late - Early deductions','59009') "
VstarConn.Execute " Insert Into NewCaptions Values('59010','Deduct Late Hours from OT','59010') "
VstarConn.Execute " Insert Into NewCaptions Values('59011','Deduct Early Hours from OT','59011') "
VstarConn.Execute " Insert Into NewCaptions Values('59012','times, total work hours of that day','59012') "
VstarConn.Execute " Insert Into NewCaptions Values('59013','OT Rule cannot be empty.','59013') "
VstarConn.Execute " Insert Into NewCaptions Values('59014','OT Rule already exists.','59014') "
VstarConn.Execute " Insert Into NewCaptions Values('59015','OT Description cannot be empty.','59015') "
VstarConn.Execute " Insert Into NewCaptions Values('59016','To timings cannot be less than Deduct timings.','59016') "
VstarConn.Execute " Insert Into NewCaptions Values('59017','Please enter To timings.','59017') "
VstarConn.Execute " Insert Into NewCaptions Values('59018','OT Rule not found.','59018') "
VstarConn.Execute " Insert Into NewCaptions Values('59019','Please Select the OT Rule.','59019') "
VstarConn.Execute " Insert Into NewCaptions Values('59020','Weekdays @','59020') "
VstarConn.Execute " Insert Into NewCaptions Values('59021','Weekoffs @','59021') "
VstarConn.Execute " Insert Into NewCaptions Values('59022','Holidays @','59022') "
VstarConn.Execute " Insert Into NewCaptions Values('59023','Deduct the following hours from the Basic OT','59023') "
VstarConn.Execute " Insert Into NewCaptions Values('59024','Basic OT will be calculated after the LATE-EARLY calculations specified in category Master','59024') "
VstarConn.Execute " Insert Into NewCaptions Values('59025','Deduct specified','59025') "
VstarConn.Execute " Insert Into NewCaptions Values('59026','--OR--','59026') "
VstarConn.Execute " Insert Into NewCaptions Values('59027','Deduct all OT','59027') "
VstarConn.Execute " Insert Into NewCaptions Values('59028','Round-Off OT','59028') "
VstarConn.Execute " Insert Into NewCaptions Values('59029','While Rounding OT only MINUTES part will be rounded , leaving the hours part as it is','59059') "
VstarConn.Execute " Insert Into NewCaptions Values('59030','Round Upto','59030') "
VstarConn.Execute " Insert Into NewCaptions Values('60001','CO Rule No.','60001') "
VstarConn.Execute " Insert Into NewCaptions Values('60002','Give CO on','60002') "
VstarConn.Execute " Insert Into NewCaptions Values('60003','CO must be availed within','60003') "
VstarConn.Execute " Insert Into NewCaptions Values('60004','minimum for 1/2 days','60004') "
VstarConn.Execute " Insert Into NewCaptions Values('60005','Minimum for Full day','60005') "
VstarConn.Execute " Insert Into NewCaptions Values('60006','CO Rule cannot be empty','60006') "
VstarConn.Execute " Insert Into NewCaptions Values('60007','CO Rule already exists','60007') "
VstarConn.Execute " Insert Into NewCaptions Values('60008','CO Description cannot be empty','60008') "
VstarConn.Execute " Insert Into NewCaptions Values('60009','Full Day Hours should be greater than Half Day Hours.','60009') "
VstarConn.Execute " Insert Into NewCaptions Values('60010','CO Rule not found','60010') "
VstarConn.Execute " Insert Into NewCaptions Values('60011','Deduct Late Hours','60011') "
VstarConn.Execute " Insert Into NewCaptions Values('60012','Deduct Early Hours','60012') "
VstarConn.Execute " Insert Into NewCaptions Values('61001','Change User Password','61001') "
VstarConn.Execute " Insert Into NewCaptions Values('61002','Old Password','61002') "
VstarConn.Execute " Insert Into NewCaptions Values('61003','New Password','61003') "
VstarConn.Execute " Insert Into NewCaptions Values('61004','Confirm Password','61004') "
VstarConn.Execute " Insert Into NewCaptions Values('61005','Passwords don''t match','61005') "
VstarConn.Execute " Insert Into NewCaptions Values('61006','Invalid User Name','61006') "
VstarConn.Execute " Insert Into NewCaptions Values('61007','Invalid Password','61007') "
VstarConn.Execute " Insert Into NewCaptions Values('62001','OT Authorization','62001') "
VstarConn.Execute " Insert Into NewCaptions Values('62002','OT Details','62003') "
VstarConn.Execute " Insert Into NewCaptions Values('62003','Transaction File not found for the Month of','62003') "
VstarConn.Execute " Insert Into NewCaptions Values('62004','Updated OT cannot be Greater than Existing OT','62004') "
VstarConn.Execute " Insert Into NewCaptions Values('62005','OT Authorized','62005') "
VstarConn.Execute " Insert Into NewCaptions Values('62006','Work Hours','62006') "
VstarConn.Execute " Insert Into NewCaptions Values('62007','OT Hrs.','62007') "
VstarConn.Execute " Insert Into NewCaptions Values('62008','No Records Found For the Employee','62008') "
VstarConn.Execute " Insert Into NewCaptions Values('62009','Cannot Update When OT hours are 0.','') "
VstarConn.Execute " Insert Into NewCaptions Values('63001','Location Master','63001') "
VstarConn.Execute " Insert Into NewCaptions Values('63002','Location Code cannot be Blank','63002') "
VstarConn.Execute " Insert Into NewCaptions Values('63003','Location Already Exists','63003') "
VstarConn.Execute " Insert Into NewCaptions Values('63004','Description cannot be Blank','63004') "
VstarConn.Execute " Insert Into NewCaptions Values('64001','Set Employee Details','64001') "
VstarConn.Execute " Insert Into NewCaptions Values('64002','Select Details','64002') "
VstarConn.Execute " Insert Into NewCaptions Values('64003','Set &' || 'Category','64003') "
VstarConn.Execute " Insert Into NewCaptions Values('64004','Set &' || 'Department','64004') "
VstarConn.Execute " Insert Into NewCaptions Values('64005','Set &' || 'Group','64005') "
VstarConn.Execute " Insert Into NewCaptions Values('64006','Set &' || 'Location','64006') "
VstarConn.Execute " Insert Into NewCaptions Values('64007','Set &' || 'OT Rule','64007') "
VstarConn.Execute " Insert Into NewCaptions Values('64008','Set CO &' || 'Rule','64008') "
VstarConn.Execute " Insert Into NewCaptions Values('64009','Set &' || 'Entries','64009') "
VstarConn.Execute " Insert Into NewCaptions Values('64010','Set Desi&' || 'gnation','64010') "
VstarConn.Execute " Insert Into NewCaptions Values('64011','Are you sure to Change the Details ?','64011') "
VstarConn.Execute " Insert Into NewCaptions Values('64012','Set Di&' || 'vision','64012') "
VstarConn.Execute " Insert Into NewCaptions Values('65001','&' || 'Back','65001') "
VstarConn.Execute " Insert Into NewCaptions Values('65002','&' || 'Next','65002') "
VstarConn.Execute " Insert Into NewCaptions Values('65003','Add','65003') "
VstarConn.Execute " Insert Into NewCaptions Values('65004','Delete','65004') "
VstarConn.Execute " Insert Into NewCaptions Values('65005','Given Below is the List of Users currently existing in the Application. The first column depicts the name of the user while the other depicts the TYPE of the user.To View /Edit the details of any user select the USER by clicking on the GRID and click Next. Thereafter just fill in the details prompted as needed. TO ADD/DELETE users click on the respective buttons.','65005') "
VstarConn.Execute " Insert Into NewCaptions Values('65006','If you are adding a New User enter the User Name of the new user you want to add. User Name can be MAXIMUM of 20 Characters. After that select the type of user you want the user to be. There are three types of user wiz Administrator, HOD i.e Head of Department and General User. The Rest of the Process will be depending heavily on the type of user selected.','65006') "
VstarConn.Execute " Insert Into NewCaptions Values('65007','Please Select the Department for which this user will work as H.O.D. It is mandatory to select the department without which further details wont be accepted and the user record will not be saved. After that select the RIGHTS for various departmental OPERATIONS to be carried out.','65007') "
VstarConn.Execute " Insert Into NewCaptions Values('65008','Please Select the Master Rights which are to be given to this User. Master Rights are the rights which are to be given on master files such as SHIFT master, GROUP master, EMPLOYEE master etc. All of the masters can be manipulated through Adding, Editing or Deleting the Records. That is why even the rights can be alloted the same way as shown below.','65008') "
VstarConn.Execute " Insert Into NewCaptions Values('65009','Please Select the Leave Transaction Rights which are to be given to this User. Leave Transaction Rights are the rights which are to be given on Leave transactions such as Opening Leave, Credit Leave , Encash Leave and Avail Leave. Leave Transactions can be either Added or Deleted, so the rights have to be given the same way as below.','65009') "
VstarConn.Execute " Insert Into NewCaptions Values('65010','Please Select the Other Rights which are given to this User. Other Rights include rights for Shift Schedule Creation, Daily Process, Monthly Process, Correction etc.  Please Check  the boxes accordingly for the rights given below.','65010') "
VstarConn.Execute " Insert Into NewCaptions Values('65011','Please enter the Passwords (Maximum 20) for this user. If the user is added as new user, old password would not be entered, but if existing user is edited, old password will be required to  change his password. Similarly user has also to enter his second level password for CRITICAL operations like Data Correction, Leave Transaction etc.','65011') "
VstarConn.Execute " Insert Into NewCaptions Values('65012','Label for Description','65012') "

VstarConn.Execute " Insert Into NewCaptions Values('65013','Old Password','65013') "
VstarConn.Execute " Insert Into NewCaptions Values('65014','New Password','65014') "
VstarConn.Execute " Insert Into NewCaptions Values('65015','Confirm Password','65015') "
VstarConn.Execute " Insert Into NewCaptions Values('65016','Leave Rights','65016') "
VstarConn.Execute " Insert Into NewCaptions Values('65017','Master Rights','65017') "
VstarConn.Execute " Insert Into NewCaptions Values('65018','Other Rights','65018') "
VstarConn.Execute " Insert Into NewCaptions Values('65019','Install','65019') "
VstarConn.Execute " Insert Into NewCaptions Values('65020','Yearly Leave Files Rights','65020') "
VstarConn.Execute " Insert Into NewCaptions Values('65021','Shift Schedule Rights','65021') "
VstarConn.Execute " Insert Into NewCaptions Values('65022','Process Rights','65022') "
VstarConn.Execute " Insert Into NewCaptions Values('65023','Daily Data Correction Rights','65023') "
VstarConn.Execute " Insert Into NewCaptions Values('65024','Report Rights','65024') "
VstarConn.Execute " Insert Into NewCaptions Values('65025','Miscellaneous Rights','65025') "
VstarConn.Execute " Insert Into NewCaptions Values('65026','Select Rights','65026') "
VstarConn.Execute " Insert Into NewCaptions Values('65027','Opening','65027') "
VstarConn.Execute " Insert Into NewCaptions Values('65028','Credit','65028') "
VstarConn.Execute " Insert Into NewCaptions Values('65029','Encash','65029') "
VstarConn.Execute " Insert Into NewCaptions Values('65030','Avail','65030') "
VstarConn.Execute " Insert Into NewCaptions Values('65031','Transactions','65031') "
VstarConn.Execute " Insert Into NewCaptions Values('65032','If This user is assigned as ADMINISTRATOR, he will automatically be assigned all the RIGHTS and PRIVILEDGES of the application.','65032') "
VstarConn.Execute " Insert Into NewCaptions Values('65033','If This user is assigned as Manager, he will given RIGHTS for the OPERATIONS pertaining to his Selection only. The Next step will be to assign a particular Rights to this user then assign him some Selection criteria.','65033') "
VstarConn.Execute " Insert Into NewCaptions Values('65034','If This user is assigned as General User, he can be RIGHTS for all the OPERATIONS in the APPLICATION except User Management, i.e he will not be able to MANIPLUATE user accounts.','65034') "
VstarConn.Execute " Insert Into NewCaptions Values('65035','User Type','65035') "
VstarConn.Execute " Insert Into NewCaptions Values('65036','[USER NAME]','65036') "
VstarConn.Execute " Insert Into NewCaptions Values('65037','User Name','65037') "
VstarConn.Execute " Insert Into NewCaptions Values('65038','[USER TYPE]','65038') "
VstarConn.Execute " Insert Into NewCaptions Values('65039','Login Password','65039') "
VstarConn.Execute " Insert Into NewCaptions Values('65040','Second Level Password','65040') "
VstarConn.Execute " Insert Into NewCaptions Values('65041','Type of Users','65041') "
VstarConn.Execute " Insert Into NewCaptions Values('65042','List of Users','65042') "
VstarConn.Execute " Insert Into NewCaptions Values('65043','User Details','65043') "
VstarConn.Execute " Insert Into NewCaptions Values('65044','HOD Details and Rights','65044') "
VstarConn.Execute " Insert Into NewCaptions Values('65045','Master Tables Rights','65045') "
VstarConn.Execute " Insert Into NewCaptions Values('65046','Leave transaction Rights','65046') "
VstarConn.Execute " Insert Into NewCaptions Values('65047','Passwords','65047') "
VstarConn.Execute " Insert Into NewCaptions Values('65048','List of Current Users','65048') "
VstarConn.Execute " Insert Into NewCaptions Values('65049','Administrator','65049') "
VstarConn.Execute " Insert Into NewCaptions Values('65050','Manager','65050') "
VstarConn.Execute " Insert Into NewCaptions Values('65051','General User','65051') "
VstarConn.Execute " Insert Into NewCaptions Values('65052','&' || 'Edit Password','65052') "
VstarConn.Execute " Insert Into NewCaptions Values('65053','Edit Parameter','65053') "
VstarConn.Execute " Insert Into NewCaptions Values('65054','On Duty','65054') "
VstarConn.Execute " Insert Into NewCaptions Values('65055','Off Duty','65055') "
VstarConn.Execute " Insert Into NewCaptions Values('65056','OT Authorization','65056') "
VstarConn.Execute " Insert Into NewCaptions Values('65057','Edit CO','65057') "
VstarConn.Execute " Insert Into NewCaptions Values('65058','Time','65058') "
VstarConn.Execute " Insert Into NewCaptions Values('65059','General Reports','65059') "
VstarConn.Execute " Insert Into NewCaptions Values('65060','View Daily Data','65060') "
VstarConn.Execute " Insert Into NewCaptions Values('65061','Reset Locks','65061') "
VstarConn.Execute " Insert Into NewCaptions Values('65062','Export Data','65062') "
VstarConn.Execute " Insert Into NewCaptions Values('65063','Create Files','65063') "
VstarConn.Execute " Insert Into NewCaptions Values('65064','Delete Old Daily Data','65064') "
VstarConn.Execute " Insert Into NewCaptions Values('65065','Edit Paid Days','65065') "
VstarConn.Execute " Insert Into NewCaptions Values('65066','Compact Database','65066') "
VstarConn.Execute " Insert Into NewCaptions Values('65067','Backup','65067') "
VstarConn.Execute " Insert Into NewCaptions Values('65068','Restore','65068') "
VstarConn.Execute " Insert Into NewCaptions Values('65069','Update Leave Balances','65069') "
VstarConn.Execute " Insert Into NewCaptions Values('65070','Create Shift Schedule','65070') "
VstarConn.Execute " Insert Into NewCaptions Values('65071','Edit Shift Shedule','65071') "
VstarConn.Execute " Insert Into NewCaptions Values('65072','Daily Process','65072') "
VstarConn.Execute " Insert Into NewCaptions Values('65073','Monthly Process','65073') "
VstarConn.Execute " Insert Into NewCaptions Values('65074','Record','65074') "
VstarConn.Execute " Insert Into NewCaptions Values('65075','Table Name','65075') "
VstarConn.Execute " Insert Into NewCaptions Values('65076','Edit','65076') "
VstarConn.Execute " Insert Into NewCaptions Values('65077','Cannot Add New User','65077') "
VstarConn.Execute " Insert Into NewCaptions Values('65078','User Cannot Delete Himself','65078') "
VstarConn.Execute " Insert Into NewCaptions Values('65079','Please Enter the User Name','65079') "
VstarConn.Execute " Insert Into NewCaptions Values('65080','User Already Exists','65080') "
VstarConn.Execute " Insert Into NewCaptions Values('65081','Please enter LOGIN Password','65081') "
VstarConn.Execute " Insert Into NewCaptions Values('65082','LOGIN Passwords don''t match','65082') "
VstarConn.Execute " Insert Into NewCaptions Values('65083','Encryption Error::LOGIN','65083') "
VstarConn.Execute " Insert Into NewCaptions Values('65084','Please enter SECOND LEVEL Password','65084') "
VstarConn.Execute " Insert Into NewCaptions Values('65085','SECOND LEVEL Passwords don''t match','65085') "
VstarConn.Execute " Insert Into NewCaptions Values('65086','Encryption Error::SECOND LEVEL','65086') "
VstarConn.Execute " Insert Into NewCaptions Values('65087','Please Select the Department for this Manager','65087') "
VstarConn.Execute " Insert Into NewCaptions Values('65088','Are You Sure to Exit','65088') "
VstarConn.Execute " Insert Into NewCaptions Values('65089','&' || 'Select/Unselect','65089') "
VstarConn.Execute " Insert Into NewCaptions Values('65090','User Accounts Management','65090') "
VstarConn.Execute " Insert Into NewCaptions Values('65091','Please select atleast One Department for the user','65091') "
VstarConn.Execute " Insert Into NewCaptions Values('65092','Please select atleast One Company for the user','65092') "
VstarConn.Execute " Insert Into NewCaptions Values('65093','Please select atleast One Group for the user','65093') "
VstarConn.Execute " Insert Into NewCaptions Values('65094','Please select atleast One Division for the user','65094') "
VstarConn.Execute " Insert Into NewCaptions Values('65095','Please select atleast One Location for the user','65095') "
VstarConn.Execute " Insert Into NewCaptions Values('65096','Given below is the List of Masters available in the system. Please select appropriate option from each of them whichever is accessible to the HOD.','65096') "
VstarConn.Execute " Insert Into NewCaptions Values('66001','Division Master','66001') "
VstarConn.Execute " Insert Into NewCaptions Values('66002','Division not Found','66002') "
VstarConn.Execute " Insert Into NewCaptions Values('66003','Division Code cannot be blank','66003') "
VstarConn.Execute " Insert Into NewCaptions Values('66004','Division Code Already Exists','66004') "
VstarConn.Execute " Insert Into NewCaptions Values('66005','Division Name cannot be blank','66005') "
VstarConn.Execute " Insert Into NewCaptions Values('67001','    Export Data','67001') "
VstarConn.Execute " Insert Into NewCaptions Values('67002','    &' || 'Next','67002') "
VstarConn.Execute " Insert Into NewCaptions Values('67003','    &' || 'Export','67003') "
VstarConn.Execute " Insert Into NewCaptions Values('67004','    &' || 'Add','67004') "
VstarConn.Execute " Insert Into NewCaptions Values('67005','    &' || 'Remove','67005') "
VstarConn.Execute " Insert Into NewCaptions Values('67006','    A&' || 'dd All','67006') "
VstarConn.Execute " Insert Into NewCaptions Values('67007','    Rem&' || 'ove All','67007') "
VstarConn.Execute " Insert Into NewCaptions Values('67008','    &' || 'Select All','67008') "
VstarConn.Execute " Insert Into NewCaptions Values('67009','    Select &' || 'Range','67009') "
VstarConn.Execute " Insert Into NewCaptions Values('67010','    U&' || 'nselect All','67010') "
VstarConn.Execute " Insert Into NewCaptions Values('67011','    &' || 'Up','67011') "
VstarConn.Execute " Insert Into NewCaptions Values('67012','    Do&' || 'wn','67012') "
VstarConn.Execute " Insert Into NewCaptions Values('67013','    Fro&' || 'm','67013') "
VstarConn.Execute " Insert Into NewCaptions Values('67014','    T&' || 'o','67014') "
VstarConn.Execute " Insert Into NewCaptions Values('67015','    Fields Available','67015') "
VstarConn.Execute " Insert Into NewCaptions Values('67016','    Select the type of Export','67016') "
VstarConn.Execute " Insert Into NewCaptions Values('67017','    Options','67017') "
VstarConn.Execute " Insert Into NewCaptions Values('67018','    Daily Data','67018') "
VstarConn.Execute " Insert Into NewCaptions Values('67019','    Monthly Data','67019') "
VstarConn.Execute " Insert Into NewCaptions Values('67020','&' || 'Back','67020') "
VstarConn.Execute " Insert Into NewCaptions Values('67021','Please Select the Fields to be Exported','67021') "
VstarConn.Execute " Insert Into NewCaptions Values('67022','Data Exported Successfully','67022') "
VstarConn.Execute " Insert Into NewCaptions Values('67023','Some Errors Occured while saving the Export Data File','67023') "
VstarConn.Execute " Insert Into NewCaptions Values('67024','Monthly Transaction file not found for the month of','67024') "
VstarConn.Execute " Insert Into NewCaptions Values('67025','Yearly Transaction file not found for the Year of','67025') "
VstarConn.Execute " Insert Into NewCaptions Values('68001','    Change Passwords','68001') "
VstarConn.Execute " Insert Into NewCaptions Values('68002','    &' || 'Change','68002') "
VstarConn.Execute " Insert Into NewCaptions Values('68003','    Old Password','68003') "
VstarConn.Execute " Insert Into NewCaptions Values('68004','    New Password','68004') "
VstarConn.Execute " Insert Into NewCaptions Values('68005','    Confirm Password','68005') "
VstarConn.Execute " Insert Into NewCaptions Values('68006','    Login Password','68006') "
VstarConn.Execute " Insert Into NewCaptions Values('68007','    Second Level Password','68007') "
VstarConn.Execute " Insert Into NewCaptions Values('68008','    Error changing password','68008') "
VstarConn.Execute " Insert Into NewCaptions Values('68009','    Please enter old password   ','68009') "
VstarConn.Execute " Insert Into NewCaptions Values('68010','    Password retreival error:: cannot continue','68010') "
VstarConn.Execute " Insert Into NewCaptions Values('68011','    User details retreival error:: cannot continue  ','68011') "
VstarConn.Execute " Insert Into NewCaptions Values('68012','    Please enter new Password','68012') "
VstarConn.Execute " Insert Into NewCaptions Values('68013','    Passwords don''t Match','68013') "
VstarConn.Execute " Insert Into NewCaptions Values('68014','    Encryption Error::Cannot Change Password','68014') "
VstarConn.Execute " Insert Into NewCaptions Values('69001','Login','69001') "
VstarConn.Execute " Insert Into NewCaptions Values('69002','User Name','69002') "
VstarConn.Execute " Insert Into NewCaptions Values('69003','Please Enter the User Name','69003') "
VstarConn.Execute " Insert Into NewCaptions Values('69004','Invalid User Name or Password','69004') "
VstarConn.Execute " Insert Into NewCaptions Values('69005','Ambiguity in HOD''s Department of','69005') "
VstarConn.Execute " Insert Into NewCaptions Values('70001','Set Shift Details for all','70001') "
VstarConn.Execute " Insert Into NewCaptions Values('70002','Set Details For','70002') "
VstarConn.Execute " Insert Into NewCaptions Values('70003','Details of Shift Info','70003') "
VstarConn.Execute " Insert Into NewCaptions Values('70004','Details of Week Off','70004') "
VstarConn.Execute " Insert Into NewCaptions Values('70005','Details of  Additional Week Off','70005') "
VstarConn.Execute " Insert Into NewCaptions Values('70006','Details Regarding Daily Process','70006') "
VstarConn.Execute " Insert Into NewCaptions Values('70007','Details Set for Selected Employees','70007') "
VstarConn.Execute " Insert Into NewCaptions Values('D0001','Employee Code','D0001') "
VstarConn.Execute " Insert Into NewCaptions Values('D0002','Shift Code','D0002') "
VstarConn.Execute " Insert Into NewCaptions Values('D0003','Dept','D0003') "
VstarConn.Execute " Insert Into NewCaptions Values('D0004','Status','D0004') "
VstarConn.Execute " Insert Into NewCaptions Values('D0005','Employee Name','D0005') "
VstarConn.Execute " Insert Into NewCaptions Values('D0006','Remarks','D0006') "
VstarConn.Execute " Insert Into NewCaptions Values('D0007','Daily absent report for the date of','D0007') "
VstarConn.Execute " Insert Into NewCaptions Values('D0008','Page','D0008') "
VstarConn.Execute " Insert Into NewCaptions Values('D0009','Date','D0009') "
VstarConn.Execute " Insert Into NewCaptions Values('D0010','Total','D0010') "
VstarConn.Execute " Insert Into NewCaptions Values('D0011','Department :','D0011') "
VstarConn.Execute " Insert Into NewCaptions Values('D0012','Category :','D0012') "
VstarConn.Execute " Insert Into NewCaptions Values('D0013','Group :','D0013') "
VstarConn.Execute " Insert Into NewCaptions Values('D0014','Arrival Time','D0014') "
VstarConn.Execute " Insert Into NewCaptions Values('D0015','Late Hours','D0015') "
VstarConn.Execute " Insert Into NewCaptions Values('D0016','Daily arrival report for the date of','D0016') "
VstarConn.Execute " Insert Into NewCaptions Values('D0017','Late Arrival Report for the date of','D0017') "
VstarConn.Execute " Insert Into NewCaptions Values('D0018','L - Late  ::  E - Early  ::  P - With Permission','D0018') "
VstarConn.Execute " Insert Into NewCaptions Values('D0019','Absent :','D0019') "
VstarConn.Execute " Insert Into NewCaptions Values('D0020','Weekly Off :','D0020') "
VstarConn.Execute " Insert Into NewCaptions Values('D0021','Leave :','D0021') "
VstarConn.Execute " Insert Into NewCaptions Values('D0022','Present :','D0022') "
VstarConn.Execute " Insert Into NewCaptions Values('D0023','Total Late','D0023') "
VstarConn.Execute " Insert Into NewCaptions Values('D0024','Category Code','D0024') "
VstarConn.Execute " Insert Into NewCaptions Values('D0025','Name of Category','D0025') "
VstarConn.Execute " Insert Into NewCaptions Values('D0026','Late Arrival Allowed','D0026') "
VstarConn.Execute " Insert Into NewCaptions Values('D0027','Early Departure Allowed','D0027') "
VstarConn.Execute " Insert Into NewCaptions Values('D0028','Late Departure Ignore','D0028') "
VstarConn.Execute " Insert Into NewCaptions Values('D0029','Early Arrival Ignore','D0029') "
VstarConn.Execute " Insert Into NewCaptions Values('D0030','Hours Required','D0030') "
VstarConn.Execute " Insert Into NewCaptions Values('D0031','Half Day Comp.Off','D0031') "
VstarConn.Execute " Insert Into NewCaptions Values('D0032','Full Day Comp. Off','D0032') "
VstarConn.Execute " Insert Into NewCaptions Values('D0033','Category Master List','D0033') "
VstarConn.Execute " Insert Into NewCaptions Values('D0034','Continuous Absent Report for the period from','D0034') "
VstarConn.Execute " Insert Into NewCaptions Values('D0035','Emp. Code','D0035') "
VstarConn.Execute " Insert Into NewCaptions Values('D0036','Emp. Name','D0036') "
VstarConn.Execute " Insert Into NewCaptions Values('D0037','Department Code','D0037') "
VstarConn.Execute " Insert Into NewCaptions Values('D0038','Name of Department','D0038') "
VstarConn.Execute " Insert Into NewCaptions Values('D0039','Department Strength','D0039') "
VstarConn.Execute " Insert Into NewCaptions Values('D0040','Department Master List','D0040') "
VstarConn.Execute " Insert Into NewCaptions Values('D0041','Daily Early Departure report for the date of','D0041') "
VstarConn.Execute " Insert Into NewCaptions Values('D0042','Dept.Time','D0042') "
VstarConn.Execute " Insert Into NewCaptions Values('D0043','Early Hours','D0043') "
VstarConn.Execute " Insert Into NewCaptions Values('D0044','Rest Out','D0044') "
VstarConn.Execute " Insert Into NewCaptions Values('D0045','Rest In','D0045') "
VstarConn.Execute " Insert Into NewCaptions Values('D0046','Work Hours','D0046') "
VstarConn.Execute " Insert Into NewCaptions Values('D0047','Extra Out','D0047') "
VstarConn.Execute " Insert Into NewCaptions Values('D0048','Extra In','D0048') "
VstarConn.Execute " Insert Into NewCaptions Values('D0049','Daily Irregular Report','D0049') "
VstarConn.Execute " Insert Into NewCaptions Values('D0050','Daily Outdoor duty report for the date of','D0050') "
VstarConn.Execute " Insert Into NewCaptions Values('D0051','OdFrm','D0051') "
VstarConn.Execute " Insert Into NewCaptions Values('D0052','OdTo','D0052') "
VstarConn.Execute " Insert Into NewCaptions Values('D0053','OT','D0053') "
VstarConn.Execute " Insert Into NewCaptions Values('D0054','Daily Overtime Report','D0054') "
VstarConn.Execute " Insert Into NewCaptions Values('D0055','Daily Performance report for the date of','D0055') "
VstarConn.Execute " Insert Into NewCaptions Values('D0056','Summary Report for the Date Of','D0056') "
VstarConn.Execute " Insert Into NewCaptions Values('D0057','Serial No','D0057') "
VstarConn.Execute " Insert Into NewCaptions Values('D0058','Total Strength','D0058') "
VstarConn.Execute " Insert Into NewCaptions Values('D0059','No.of Emp.','D0059') "
VstarConn.Execute " Insert Into NewCaptions Values('D0060','OD','D0060') "

VstarConn.Execute " Insert Into NewCaptions Values('D0061','Employee Master Details','D0061') "
VstarConn.Execute " Insert Into NewCaptions Values('D0062','Card No','D0062') "
VstarConn.Execute " Insert Into NewCaptions Values('D0063','Designation','D0063') "
VstarConn.Execute " Insert Into NewCaptions Values('D0064','Shift Type','D0064') "
VstarConn.Execute " Insert Into NewCaptions Values('D0065','Joining Date','D0065') "
VstarConn.Execute " Insert Into NewCaptions Values('D0066','Min.entries','D0066') "
VstarConn.Execute " Insert Into NewCaptions Values('D0067','Basic Salary','D0067') "
VstarConn.Execute " Insert Into NewCaptions Values('D0068','Current Address','D0068') "
VstarConn.Execute " Insert Into NewCaptions Values('D0069','Address','D0069') "
VstarConn.Execute " Insert Into NewCaptions Values('D0070','City','D0070') "
VstarConn.Execute " Insert Into NewCaptions Values('D0071','Tel','D0071') "
VstarConn.Execute " Insert Into NewCaptions Values('D0072','PinCode','D0072') "
VstarConn.Execute " Insert Into NewCaptions Values('D0073','Sex','D0073') "
VstarConn.Execute " Insert Into NewCaptions Values('D0074','Blood group','D0074') "
VstarConn.Execute " Insert Into NewCaptions Values('D0075','Birth date','D0075') "
VstarConn.Execute " Insert Into NewCaptions Values('D0076','Permanent Address','D0076') "
VstarConn.Execute " Insert Into NewCaptions Values('D0077','District','D0077') "
VstarConn.Execute " Insert Into NewCaptions Values('D0078','State','D0078') "
VstarConn.Execute " Insert Into NewCaptions Values('D0079','PhoneNo','D0079') "
VstarConn.Execute " Insert Into NewCaptions Values('D0080','Remarks Comments','D0080') "
VstarConn.Execute " Insert Into NewCaptions Values('D0081','Left Date','D0081') "
VstarConn.Execute " Insert Into NewCaptions Values('D0082','Daily entry report for','D0082') "
VstarConn.Execute " Insert Into NewCaptions Values('D0083','Punches','D0083') "
VstarConn.Execute " Insert Into NewCaptions Values('D0084','Group Master','D0084') "
VstarConn.Execute " Insert Into NewCaptions Values('D0085','Group Code','D0085') "
VstarConn.Execute " Insert Into NewCaptions Values('D0086','Group Description','D0086') "
VstarConn.Execute " Insert Into NewCaptions Values('D0087','Holiday Master List','D0087') "
VstarConn.Execute " Insert Into NewCaptions Values('D0088','Holiday Date','D0088') "
VstarConn.Execute " Insert Into NewCaptions Values('D0089','Holiday Description','D0089') "
VstarConn.Execute " Insert Into NewCaptions Values('D0090','Leave Master List','D0090') "
VstarConn.Execute " Insert Into NewCaptions Values('D0091','Paid Leave','D0091') "
VstarConn.Execute " Insert Into NewCaptions Values('D0092','Balance','D0092') "
VstarConn.Execute " Insert Into NewCaptions Values('D0093','Encash','D0093') "
VstarConn.Execute " Insert Into NewCaptions Values('D0094','Leave days calculation','D0094') "
VstarConn.Execute " Insert Into NewCaptions Values('D0095','Yearly','D0095') "
VstarConn.Execute " Insert Into NewCaptions Values('D0096','No of Leave','D0096') "
VstarConn.Execute " Insert Into NewCaptions Values('D0097','Accumulation','D0097') "
VstarConn.Execute " Insert Into NewCaptions Values('D0098','Employee Master List','D0098') "
VstarConn.Execute " Insert Into NewCaptions Values('D0099','Daily Manpower report for the date of','D0099') "
VstarConn.Execute " Insert Into NewCaptions Values('D0100','Present','D0100') "
VstarConn.Execute " Insert Into NewCaptions Values('D0101','Absent','D0101') "
VstarConn.Execute " Insert Into NewCaptions Values('D0102','Offs','D0102') "
VstarConn.Execute " Insert Into NewCaptions Values('D0103','Monthly Performance report for the month of','D0103') "
VstarConn.Execute " Insert Into NewCaptions Values('D0104','Overtime report for the month of','D0104') "
VstarConn.Execute " Insert Into NewCaptions Values('D0105','L A T E S','D0105') "
VstarConn.Execute " Insert Into NewCaptions Values('D0106','E A R L Y','D0106') "
VstarConn.Execute " Insert Into NewCaptions Values('D0107','No.','D0107') "
VstarConn.Execute " Insert Into NewCaptions Values('D0108','Absent / Late / Early Summary report for the month of','D0108') "
VstarConn.Execute " Insert Into NewCaptions Values('D0109','Monthly attendance for the month of','D0109') "
VstarConn.Execute " Insert Into NewCaptions Values('D0110','OT_HRS','D0110') "
VstarConn.Execute " Insert Into NewCaptions Values('D0111','Total Early','D0111') "
VstarConn.Execute " Insert Into NewCaptions Values('D0112','Monthly Late Arrival Report for the month of','D0112') "
VstarConn.Execute " Insert Into NewCaptions Values('D0113','Monthly Early Departure Report for the month of','D0113') "
VstarConn.Execute " Insert Into NewCaptions Values('D0114','Leave Balance report for the month of','D0114') "
VstarConn.Execute " Insert Into NewCaptions Values('D0115','Leave Code','D0115') "
VstarConn.Execute " Insert Into NewCaptions Values('D0116','Leave Name','D0116') "
VstarConn.Execute " Insert Into NewCaptions Values('D0117','Leave From','D0117') "
VstarConn.Execute " Insert Into NewCaptions Values('D0118','Leave To','D0118') "
VstarConn.Execute " Insert Into NewCaptions Values('D0119','Leave Days','D0119') "
VstarConn.Execute " Insert Into NewCaptions Values('D0120','Time','D0120') "
VstarConn.Execute " Insert Into NewCaptions Values('D0121','2nd Punch','D0121') "
VstarConn.Execute " Insert Into NewCaptions Values('D0122','3rd Punch','D0122') "
VstarConn.Execute " Insert Into NewCaptions Values('D0123','4th Punch','D0123') "
VstarConn.Execute " Insert Into NewCaptions Values('D0124','5th Punch','D0124') "
VstarConn.Execute " Insert Into NewCaptions Values('D0125','Leave Availed for the Month of','D0125') "
VstarConn.Execute " Insert Into NewCaptions Values('D0190','Early Departure memo for the month of','D0190') "
VstarConn.Execute " Insert Into NewCaptions Values('D0191','Late Arrival Memo for the month of','D0191') "
VstarConn.Execute " Insert Into NewCaptions Values('D0192','Absent Memo for the month of','D0192') "
VstarConn.Execute " Insert Into NewCaptions Values('D0193','O.T. Paid','D0193') "
VstarConn.Execute " Insert Into NewCaptions Values('D0194','O.T. Done','D0194') "
VstarConn.Execute " Insert Into NewCaptions Values('D0195','Overtime Paid report for the month of','D0195') "
VstarConn.Execute " Insert Into NewCaptions Values('D0196','Monthly Absent  Report for the month of','D0196') "
VstarConn.Execute " Insert Into NewCaptions Values('D0197','Monthly Present Report for the month of','D0197') "
VstarConn.Execute " Insert Into NewCaptions Values('D0198','Monthly muster for the month of','D0198') "
VstarConn.Execute " Insert Into NewCaptions Values('D0199','Monthly shift schedule report for the month of','D0199') "
VstarConn.Execute " Insert Into NewCaptions Values('D0201','Yearly Performance Report for the year','D0201') "
VstarConn.Execute " Insert Into NewCaptions Values('D0202','PAIDDAYS','D0202') "
VstarConn.Execute " Insert Into NewCaptions Values('D0203','LT_NO','D0203') "
VstarConn.Execute " Insert Into NewCaptions Values('D0204','LT_HRS','D0204') "
VstarConn.Execute " Insert Into NewCaptions Values('D0205','ERL_NO','D0205') "
VstarConn.Execute " Insert Into NewCaptions Values('D0206','ERL_HRS','D0206') "
VstarConn.Execute " Insert Into NewCaptions Values('D0207','WRK_HRS','D0207') "
VstarConn.Execute " Insert Into NewCaptions Values('D0208','Yearly Man-Power Utilisation Report for the year','D0208') "
VstarConn.Execute " Insert Into NewCaptions Values('D0209','NIGHT','D0209') "
VstarConn.Execute " Insert Into NewCaptions Values('D0210','Leave Utilisation Report for year','D0210') "
VstarConn.Execute " Insert Into NewCaptions Values('D0211','Leave Code','D0211') "
VstarConn.Execute " Insert Into NewCaptions Values('D0212','Remarks / Period','D0212') "
VstarConn.Execute " Insert Into NewCaptions Values('D0213','Credited','D0213') "
VstarConn.Execute " Insert Into NewCaptions Values('D0214','Availed','D0214') "
VstarConn.Execute " Insert Into NewCaptions Values('D0215','Yearly Absent Report for the year','D0215') "
VstarConn.Execute " Insert Into NewCaptions Values('D0216','Yearly Present Report for the year','D0216') "
VstarConn.Execute " Insert Into NewCaptions Values('D0217','Weekly Performance report for the period of','D0217') "
VstarConn.Execute " Insert Into NewCaptions Values('D0218','To','D0218') "
VstarConn.Execute " Insert Into NewCaptions Values('D0219','Arr','D0219') "
VstarConn.Execute " Insert Into NewCaptions Values('D0220','Dep','D0220') "
VstarConn.Execute " Insert Into NewCaptions Values('D0221','Late','D0221') "
VstarConn.Execute " Insert Into NewCaptions Values('D0222','Earl','D0222') "
VstarConn.Execute " Insert Into NewCaptions Values('D0223','Work','D0223') "
VstarConn.Execute " Insert Into NewCaptions Values('D0224','Rem','D0224') "
VstarConn.Execute " Insert Into NewCaptions Values('D0225','Shf','D0225') "
VstarConn.Execute " Insert Into NewCaptions Values('D0226','Weekly Absent report for the period of','D0226') "
VstarConn.Execute " Insert Into NewCaptions Values('D0227','Weekly Attendance report for the period of','D0227') "
VstarConn.Execute " Insert Into NewCaptions Values('D0228','Weekly Early Departure report for the period of','D0228') "
VstarConn.Execute " Insert Into NewCaptions Values('D0229','Weekly Late Arrival report for the period of','D0229') "
VstarConn.Execute " Insert Into NewCaptions Values('D0230','Weekly Overtime report for the period of','D0230') "
VstarConn.Execute " Insert Into NewCaptions Values('D0231','Weekly Shift Arrangement report for the period of','D0231') "
VstarConn.Execute " Insert Into NewCaptions Values('D0232','Weekly Irregular Punch report from','D0232') "
VstarConn.Execute " Insert Into NewCaptions Values('D0233','Shift Master List','D0233') "
VstarConn.Execute " Insert Into NewCaptions Values('D0234','Name of the shift','D0234') "
VstarConn.Execute " Insert Into NewCaptions Values('D0235','Shift Time','D0235') "
VstarConn.Execute " Insert Into NewCaptions Values('D0236','Starting','D0236') "
VstarConn.Execute " Insert Into NewCaptions Values('D0237','Ending','D0237') "
VstarConn.Execute " Insert Into NewCaptions Values('D0238','Hours','D0238') "
VstarConn.Execute " Insert Into NewCaptions Values('D0239','Lunch Time','D0239') "
VstarConn.Execute " Insert Into NewCaptions Values('D0240','Half Day Time','D0240') "
VstarConn.Execute " Insert Into NewCaptions Values('D0241','Daily Shift arrangement report for the date of','D0241') "
VstarConn.Execute " Insert Into NewCaptions Values('D0242','Rotational Shift Master List','D0242') "
VstarConn.Execute " Insert Into NewCaptions Values('D0243','Rotation Code','D0243') "
VstarConn.Execute " Insert Into NewCaptions Values('D0244','Type of Rotation','D0244') "
VstarConn.Execute " Insert Into NewCaptions Values('D0245','Shift Pattern','D0245') "
VstarConn.Execute " Insert Into NewCaptions Values('D0246','Rotation Pattern','D0246') "
VstarConn.Execute " Insert Into NewCaptions Values('D0247','Overtime Report for the Period from','D0247') "
VstarConn.Execute " Insert Into NewCaptions Values('D0248','Performance report from','D0248') "
VstarConn.Execute " Insert Into NewCaptions Values('D0249','Attendance Muster report for the period from','D0249') "
VstarConn.Execute " Insert Into NewCaptions Values('D0250','Late Arrival report for the period from','D0250') "
VstarConn.Execute " Insert Into NewCaptions Values('D0251','Early Departure report for the period from','D0251') "
VstarConn.Execute " Insert Into NewCaptions Values('D0252','WO on holiday report for the month of','D0252') "
VstarConn.Execute " Insert Into NewCaptions Values('D0253','Days to be deducted','D0253') "
VstarConn.Execute " Insert Into NewCaptions Values('D0254','Total Lates report for the month of','D0254') "
VstarConn.Execute " Insert Into NewCaptions Values('D0255','Total Early report for the month of','D0255') "
VstarConn.Execute " Insert Into NewCaptions Values('D0256','Division Master','D0256') "
VstarConn.Execute " Insert Into NewCaptions Values('D0257','Strength','D0257') "
VstarConn.Execute " Insert Into NewCaptions Values('D0258','M 1','D0258') "
VstarConn.Execute " Insert Into NewCaptions Values('D0259','M 2','D0259') "
VstarConn.Execute " Insert Into NewCaptions Values('D0260','Meal Allowance Report From','D0260') "
VstarConn.Execute " Insert Into NewCaptions Values('D0261','Summary report from','D0261') "
VstarConn.Execute " Insert Into NewCaptions Values('D0262','Permission card report From','D0262') "
VstarConn.Execute " Insert Into NewCaptions Values('D0263','Leave Availment Report From ','D0263') "
VstarConn.Execute " Insert Into NewCaptions Values('D0264','Corrected','D0264') "
VstarConn.Execute " Insert Into NewCaptions Values('D0265 ','C','D0265') "
VstarConn.Execute " Insert Into NewCaptions Values('M1001','Invalid Date Format Found:: Cannot Proceed','M1001') "
VstarConn.Execute " Insert Into NewCaptions Values('M1002','Please set Your Date Settings to ''M/D/YY'' Format in','M1002') "
VstarConn.Execute " Insert Into NewCaptions Values('M1003','Start --> Settings --> Control Panel --> Regional Settings --> Date','M1003') "
VstarConn.Execute " Insert Into NewCaptions Values('M1004','--> Short Date Style','M1004') "
VstarConn.Execute " Insert Into NewCaptions Values('M1005','Please set Your Date Settings to ''DD/MM/YY'' Format in','M1005') "
VstarConn.Execute " Insert Into NewCaptions Values('M1006','Your Regional Date Settings do not Match the Application Date Settings','M1006') "
VstarConn.Execute " Insert Into NewCaptions Values('M1007','(American Type)','M1007') "
VstarConn.Execute " Insert Into NewCaptions Values('M1008','Do You Wish to Set It','M1008') "
VstarConn.Execute " Insert Into NewCaptions Values('M1009','( British Type)','M1009') "
VstarConn.Execute " Insert Into NewCaptions Values('M1010','One or More Required Script Files are Missing :: Cannot Run Daily Process.','M1010') "
VstarConn.Execute " Insert Into NewCaptions Values('M1011','Total Calculation File Needed for Monthly Processing not Found','M1011') "
VstarConn.Execute " Insert Into NewCaptions Values('M1012','Rules File Needed for Monthly Processing not Found','M1012') "
VstarConn.Execute " Insert Into NewCaptions Values('M1013','Error Loading Total Calculation File Needed for Monthly Processing','M1013') "
VstarConn.Execute " Insert Into NewCaptions Values('M1014','Error Loading Rules File Needed for Monthly Processing','M1014') "
VstarConn.Execute " Insert Into NewCaptions Values('M1015','[Demo Version]','M1015') "
VstarConn.Execute " Insert Into NewCaptions Values('M3001','CO not Found :: Leave Balance File for the Current Year not Updated','M3001') "
VstarConn.Execute " Insert Into NewCaptions Values('M3002','CO not Found in Leave Master','M3002') "
VstarConn.Execute " Insert Into NewCaptions Values('M3003','Leave Balanace File for the Current Year not Found','M3003') "
VstarConn.Execute " Insert Into NewCaptions Values('M3004','Please Create it First and then do the Daily Process','M3004') "
VstarConn.Execute " Insert Into NewCaptions Values('M3005','Reading from File','M3005') "
VstarConn.Execute " Insert Into NewCaptions Values('M4001','Month should be between 01 to 12.','M4001') "
VstarConn.Execute " Insert Into NewCaptions Values('M4002','Invalid Number of Days','M4002') "
VstarConn.Execute " Insert Into NewCaptions Values('M4003','Invalid Date Length','M4003') "
VstarConn.Execute " Insert Into NewCaptions Values('M4004','Invalid Date Structure','M4004') "
VstarConn.Execute " Insert Into NewCaptions Values('M6001','Yearly Leave Files are not Created :: Please Create them','M6001') "
VstarConn.Execute " Insert Into NewCaptions Values('M6002','Yearly Leave Files for the Year','M6002') "
VstarConn.Execute " Insert Into NewCaptions Values('M6003','Yearly Updation Aborted','M6003') "
VstarConn.Execute " Insert Into NewCaptions Values('M6004','Yearly Leaves Files Updated','M6004') "
VstarConn.Execute " Insert Into NewCaptions Values('M6005','Yearly Files are not Created Properly :: Please Re-Create it.','M6005') "
VstarConn.Execute " Insert Into NewCaptions Values('M7001','Shift File for the Month Of','M7001') "
VstarConn.Execute " Insert Into NewCaptions Values('M7002','Please Select the required Report.','M7002') "
VstarConn.Execute " Insert Into NewCaptions Values('M7003','Invalid File Name','M7003') "
VstarConn.Execute " Insert Into NewCaptions Values('M7004',' saved successfully','M7004') "
VstarConn.Execute " Insert Into NewCaptions Values('M7005','Yearly Leave Transaction File Not Found','M7005') "
VstarConn.Execute " Insert Into NewCaptions Values('M7006','Mail','M7006') "
VstarConn.Execute " Insert Into NewCaptions Values('M7007','View','M7007') "
VstarConn.Execute " Insert Into NewCaptions Values('M7008','Print','M7008') "
VstarConn.Execute " Insert Into NewCaptions Values('M7009','Print to File','M7009') "
VstarConn.Execute " Insert Into NewCaptions Values('M7010','   Checking Validations ..','M7010') "
VstarConn.Execute " Insert Into NewCaptions Values('M7011','   Processing Valid data ..','M7011') "
VstarConn.Execute " Insert Into NewCaptions Values('M7012','   Executing Query ..','M7012') "
VstarConn.Execute " Insert Into NewCaptions Values('M7013','   Operation Aborted','M7013') "
VstarConn.Execute " Insert Into NewCaptions Values('M7014','   Preparing Report to','M7014') "
VstarConn.Execute " Insert Into NewCaptions Values('M7015','   Daily Reports','M7015') "
VstarConn.Execute " Insert Into NewCaptions Values('M7016','   Weekly Reports','M7016') "
VstarConn.Execute " Insert Into NewCaptions Values('M7017','   Monthly Reports','M7017') "
VstarConn.Execute " Insert Into NewCaptions Values('M7018','   Yearly Reports','M7018') "
VstarConn.Execute " Insert Into NewCaptions Values('M7019','   Masters Reports','M7019') "
VstarConn.Execute " Insert Into NewCaptions Values('M7020','   Periodic Reports','M7020') "
VstarConn.Execute " Insert Into NewCaptions Values('M7021','    Sending Mail to','M7021') "
VstarConn.Execute " Insert Into NewCaptions Values('M7022','Opening','M7022') "
VstarConn.Execute " Insert Into NewCaptions Values('M7023','Credited','M7023') "
VstarConn.Execute " Insert Into NewCaptions Values('M7024','Encashed','M7024') "
VstarConn.Execute " Insert Into NewCaptions Values('M7025','Late Cut','M7025') "
VstarConn.Execute " Insert Into NewCaptions Values('M7026','Early Cut','M7026') "

VstarConn.Execute "COMMIT"
InsertOracleCaptions = True
Exit Function
Err_P:
    MsgBox Err.Description
    Resume Next
End Function

Public Function InsertSQLCaptions() As Boolean
On Error GoTo Err_P
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute "Delete from Newcaptions"
VstarConn.Execute " Insert Into NewCaptions Values('00001','Not Enough Rights, Access Denied','00001') "
VstarConn.Execute " Insert Into NewCaptions Values('00002','&OK','00002') "
VstarConn.Execute " Insert Into NewCaptions Values('00003','&Cancel','00003') "
VstarConn.Execute " Insert Into NewCaptions Values('00004','&Add','00004') "
VstarConn.Execute " Insert Into NewCaptions Values('00005','&Edit','00005') "
VstarConn.Execute " Insert Into NewCaptions Values('00006','&Delete','00006') "
VstarConn.Execute " Insert Into NewCaptions Values('00007','&Save','00007') "
VstarConn.Execute " Insert Into NewCaptions Values('00008','E&xit','00008') "
VstarConn.Execute " Insert Into NewCaptions Values('00009','Continue ?','00009') "
VstarConn.Execute " Insert Into NewCaptions Values('00010','From','00010') "
VstarConn.Execute " Insert Into NewCaptions Values('00011','To','00011') "
VstarConn.Execute " Insert Into NewCaptions Values('00012','Days','00012') "
VstarConn.Execute " Insert Into NewCaptions Values('00013','List','00013') "
VstarConn.Execute " Insert Into NewCaptions Values('00014','Details','00014') "
VstarConn.Execute " Insert Into NewCaptions Values('00015','Are you sure to Delete this Record ?','00015') "
VstarConn.Execute " Insert Into NewCaptions Values('00016','From Date cannot be Empty.','00016') "
VstarConn.Execute " Insert Into NewCaptions Values('00017','To Date cannot be Empty.','00017') "
VstarConn.Execute " Insert Into NewCaptions Values('00018','From Date cannot be greater than To Date.','00018') "
VstarConn.Execute " Insert Into NewCaptions Values('00019','From Date','00019') "
VstarConn.Execute " Insert Into NewCaptions Values('00020','To Date','00020') "
VstarConn.Execute " Insert Into NewCaptions Values('00021',' not in Current Year.','00021') "
VstarConn.Execute " Insert Into NewCaptions Values('00022','&Close','00022') "
VstarConn.Execute " Insert Into NewCaptions Values('00023','Hours','00023') "
VstarConn.Execute " Insert Into NewCaptions Values('00024','Minutes cannot be greater than 0.59','00024') "
VstarConn.Execute " Insert Into NewCaptions Values('00025','Maximum value cannot be greater than 23.59','00025') "
VstarConn.Execute " Insert Into NewCaptions Values('00026','Month','00026') "
VstarConn.Execute " Insert Into NewCaptions Values('00027','Year','00027') "
VstarConn.Execute " Insert Into NewCaptions Values('00028','&Month','00028') "
VstarConn.Execute " Insert Into NewCaptions Values('00029','&Year','00029') "
VstarConn.Execute " Insert Into NewCaptions Values('00030','Date','00030') "
VstarConn.Execute " Insert Into NewCaptions Values('00031','Shift','00031') "
VstarConn.Execute " Insert Into NewCaptions Values('00032','Entry','00032') "
VstarConn.Execute " Insert Into NewCaptions Values('00033','Status','00033') "
VstarConn.Execute " Insert Into NewCaptions Values('00034','Arrival','00034') "
VstarConn.Execute " Insert Into NewCaptions Values('00035','Late','00035') "
VstarConn.Execute " Insert Into NewCaptions Values('00036','Departure','00036') "
VstarConn.Execute " Insert Into NewCaptions Values('00037','Early','00037') "
VstarConn.Execute " Insert Into NewCaptions Values('00038','Overtime','00038') "
VstarConn.Execute " Insert Into NewCaptions Values('00039','&Finish','00039') "
VstarConn.Execute " Insert Into NewCaptions Values('00040','Select &Range','00040') "
VstarConn.Execute " Insert Into NewCaptions Values('00041','&Unselect Range','00041') "
VstarConn.Execute " Insert Into NewCaptions Values('00042','&Select All','00042') "
VstarConn.Execute " Insert Into NewCaptions Values('00043','U&nselect All','00043') "
VstarConn.Execute " Insert Into NewCaptions Values('00044','Employee Selection','00044') "
VstarConn.Execute " Insert Into NewCaptions Values('00045','Fro&m','00045') "
VstarConn.Execute " Insert Into NewCaptions Values('00046','T&o','00046') "
VstarConn.Execute " Insert Into NewCaptions Values('00047','Code','00047') "
VstarConn.Execute " Insert Into NewCaptions Values('00048','Name','00048') "
VstarConn.Execute " Insert Into NewCaptions Values('00049','Please Select the Employees','00049') "
VstarConn.Execute " Insert Into NewCaptions Values('00050','Please Select the Dat File.','00050') "
VstarConn.Execute " Insert Into NewCaptions Values('00051','Category','00051') "
VstarConn.Execute " Insert Into NewCaptions Values('00052','Description','00052') "
VstarConn.Execute " Insert Into NewCaptions Values('00053','&Create','00053') "
VstarConn.Execute " Insert Into NewCaptions Values('00054','Leave Transaction File for the year.','00054') "
VstarConn.Execute " Insert Into NewCaptions Values('00055',' not found.','00055') "
VstarConn.Execute " Insert Into NewCaptions Values('00056','Decimal value can be 0.5 or 0 Only','00056') "
VstarConn.Execute " Insert Into NewCaptions Values('00057','Company','00057') "
VstarConn.Execute " Insert Into NewCaptions Values('00058','Department','00058') "
VstarConn.Execute " Insert Into NewCaptions Values('00059','Group','00059') "
VstarConn.Execute " Insert Into NewCaptions Values('00060','Minimum Value cannot be less than 0.','00060') "
VstarConn.Execute " Insert Into NewCaptions Values('00061','Employee Code','00061') "
VstarConn.Execute " Insert Into NewCaptions Values('00062','Data is  in Process : Please Try after Some Time','00062') "
VstarConn.Execute " Insert Into NewCaptions Values('00063','Leave','00063') "
VstarConn.Execute " Insert Into NewCaptions Values('00064','&Reset','00064') "
VstarConn.Execute " Insert Into NewCaptions Values('00065','Mon','00065') "
VstarConn.Execute " Insert Into NewCaptions Values('00066','Tue','00066') "
VstarConn.Execute " Insert Into NewCaptions Values('00067','Wed','00067') "
VstarConn.Execute " Insert Into NewCaptions Values('00068','Thu','00068') "
VstarConn.Execute " Insert Into NewCaptions Values('00069','Fri','00069') "
VstarConn.Execute " Insert Into NewCaptions Values('00070','Sat','00070') "
VstarConn.Execute " Insert Into NewCaptions Values('00071','Sun','00071') "
VstarConn.Execute " Insert Into NewCaptions Values('00072','Date cannot be blank','00072') "
VstarConn.Execute " Insert Into NewCaptions Values('00073','Non-Leap Year cannot have 29 days in February','00073') "
VstarConn.Execute " Insert Into NewCaptions Values('00074','Selec&t Printer','00074') "
VstarConn.Execute " Insert Into NewCaptions Values('00075','&Send','00075') "
VstarConn.Execute " Insert Into NewCaptions Values('00076','Pre&view','00076') "
VstarConn.Execute " Insert Into NewCaptions Values('00077','&Print','00077') "
VstarConn.Execute " Insert Into NewCaptions Values('00078','&File','00078') "
VstarConn.Execute " Insert Into NewCaptions Values('00079','No Records found.','00079') "
VstarConn.Execute " Insert Into NewCaptions Values('00080','You have been marked absent on following dates. Please fill up OD/LEAVES at the earliest','00080') "
VstarConn.Execute " Insert Into NewCaptions Values('00081','You have been marked late on the following dates. Kindly forward OD/REQUISITE PERMISSION within 3 days','00081') "
VstarConn.Execute " Insert Into NewCaptions Values('00082','You have been marked early on following dates. Kindly forward ON/LEAVES/REQUISITE PERMISSION at the earliest','00082') "
VstarConn.Execute " Insert Into NewCaptions Values('00083','Login','00083') "
VstarConn.Execute " Insert Into NewCaptions Values('00084','Password','00084') "
VstarConn.Execute " Insert Into NewCaptions Values('00085','Invalid User','00085') "
VstarConn.Execute " Insert Into NewCaptions Values('00086','Incorrect PassWord','00086') "
VstarConn.Execute " Insert Into NewCaptions Values('00087','One or More Required Yearly Leave Files Missing :: Please Re-Create','00087') "
VstarConn.Execute " Insert Into NewCaptions Values('00088','OT Rules','00088') "
VstarConn.Execute " Insert Into NewCaptions Values('00089','CO Rules','00089') "
VstarConn.Execute " Insert Into NewCaptions Values('00090','OT Rule','00090') "
VstarConn.Execute " Insert Into NewCaptions Values('00091','CO Rule','00091') "
VstarConn.Execute " Insert Into NewCaptions Values('00092','Weekdays','00092') "
VstarConn.Execute " Insert Into NewCaptions Values('00093','Weekoffs','00093') "
VstarConn.Execute " Insert Into NewCaptions Values('00094','Holidays','00094') "
VstarConn.Execute " Insert Into NewCaptions Values('00095','Present','00095') "
VstarConn.Execute " Insert Into NewCaptions Values('00096','Absent','00096') "
VstarConn.Execute " Insert Into NewCaptions Values('00097','Deductions','00097') "
VstarConn.Execute " Insert Into NewCaptions Values('00098','Deduct','00098') "
VstarConn.Execute " Insert Into NewCaptions Values('00099','All','00099') "
VstarConn.Execute " Insert Into NewCaptions Values('00100','YES','00100') "
VstarConn.Execute " Insert Into NewCaptions Values('00101','NO','00101') "
VstarConn.Execute " Insert Into NewCaptions Values('00102','Invalid User Name or Password','00102') "
VstarConn.Execute " Insert Into NewCaptions Values('00103','Password changed successfully','00103') "
VstarConn.Execute " Insert Into NewCaptions Values('00104','OT Authorization','00104') "
VstarConn.Execute " Insert Into NewCaptions Values('00105','Reports','00105') "
VstarConn.Execute " Insert Into NewCaptions Values('00106','Report','00106') "
VstarConn.Execute " Insert Into NewCaptions Values('00107','Unacceptable Password :: Please Type New Password','00107') "
VstarConn.Execute " Insert Into NewCaptions Values('00108','Please Enter Password','00108') "
VstarConn.Execute " Insert Into NewCaptions Values('00109','Locations','00109') "
VstarConn.Execute " Insert Into NewCaptions Values('00110','Location','00110') "
VstarConn.Execute " Insert Into NewCaptions Values('00111','Click to Toggle Selection','00111') "
VstarConn.Execute " Insert Into NewCaptions Values('00112','Operation Invalid before Employee Joindate','00112') "
VstarConn.Execute " Insert Into NewCaptions Values('00113','To timings cannot be less then from timings','00113') "
VstarConn.Execute " Insert Into NewCaptions Values('00114','&Abort','00114') "
VstarConn.Execute " Insert Into NewCaptions Values('00115','Are you Sure to Abort the Process ?','00115') "
VstarConn.Execute " Insert Into NewCaptions Values('00116','Unauthorized OT','00116') "
VstarConn.Execute " Insert Into NewCaptions Values('00117','Authorized','00117') "
VstarConn.Execute " Insert Into NewCaptions Values('00118','Unauthorized','00118') "
VstarConn.Execute " Insert Into NewCaptions Values('00119','Authorized OT','00119') "
VstarConn.Execute " Insert Into NewCaptions Values('00120','Father Name','00120') "
VstarConn.Execute " Insert Into NewCaptions Values('00121','Entry','00121') "
VstarConn.Execute " Insert Into NewCaptions Values('00122','Entries','00122') "
VstarConn.Execute " Insert Into NewCaptions Values('00123','Designation','00123') "
VstarConn.Execute " Insert Into NewCaptions Values('00124','Remarks','00124') "
VstarConn.Execute " Insert Into NewCaptions Values('00125','OT Remark','00125') "
VstarConn.Execute " Insert Into NewCaptions Values('00126','Division','00126') "
VstarConn.Execute " Insert Into NewCaptions Values('00127','Divisions','00127') "
VstarConn.Execute " Insert Into NewCaptions Values('00128','Referential Integrity Error. Cannot Delete this record.','00128') "
VstarConn.Execute " Insert Into NewCaptions Values('01001','Select Date','01001') "
VstarConn.Execute " Insert Into NewCaptions Values('01002','&OK','01002') "
VstarConn.Execute " Insert Into NewCaptions Values('01003','&Cancel','01003') "
VstarConn.Execute " Insert Into NewCaptions Values('02001','Change Password','02001') "
VstarConn.Execute " Insert Into NewCaptions Values('02002','Old Settings','02002') "
VstarConn.Execute " Insert Into NewCaptions Values('02003','Menu Item','02003') "
VstarConn.Execute " Insert Into NewCaptions Values('02004','Old Password','02004') "
VstarConn.Execute " Insert Into NewCaptions Values('02005','Settings','02005') "
VstarConn.Execute " Insert Into NewCaptions Values('02006','New Password','02006') "
VstarConn.Execute " Insert Into NewCaptions Values('02007','Reconfirm Password','02007') "
VstarConn.Execute " Insert Into NewCaptions Values('02008','Please Select the Option for which the Password is to be Changed','02008') "
VstarConn.Execute " Insert Into NewCaptions Values('02009','Please Enter Old Password','02009') "
VstarConn.Execute " Insert Into NewCaptions Values('02010','Please Enter New Password','02010') "
VstarConn.Execute " Insert Into NewCaptions Values('02011','Please Confirm New Password','02011') "
VstarConn.Execute " Insert Into NewCaptions Values('02012','Cannot confirm new Password. Please try again.','02012') "
VstarConn.Execute " Insert Into NewCaptions Values('02013','Invalid Password.','02013') "
VstarConn.Execute " Insert Into NewCaptions Values('02014','Unacceptable Password :: Please Type New Password','02014') "
VstarConn.Execute " Insert Into NewCaptions Values('02015','Password for','02015') "
VstarConn.Execute " Insert Into NewCaptions Values('02016',' Changed Successfully','02016') "
VstarConn.Execute " Insert Into NewCaptions Values('03001','Select Path for .DAT File','03001') "
VstarConn.Execute " Insert Into NewCaptions Values('03002','Select Directory','03002') "
VstarConn.Execute " Insert Into NewCaptions Values('03003','Drive:','03003') "
VstarConn.Execute " Insert Into NewCaptions Values('03004','Select','03004') "
VstarConn.Execute " Insert Into NewCaptions Values('03005','Cancel','03005') "
VstarConn.Execute " Insert Into NewCaptions Values('04001','Send Report to. . .','04001') "
VstarConn.Execute " Insert Into NewCaptions Values('04002','Type Subject here:','04002') "
VstarConn.Execute " Insert Into NewCaptions Values('04003',' Send Report to whom?','04003') "
VstarConn.Execute " Insert Into NewCaptions Values('04004','Send to Manager','04004') "
VstarConn.Execute " Insert Into NewCaptions Values('04005','Send to each Employee','04005') "
VstarConn.Execute " Insert Into NewCaptions Values('05001','About Visual Star(VSTAR)','05001') "
VstarConn.Execute " Insert Into NewCaptions Values('05002','Visual System for Time Attendance Recording','05002') "
VstarConn.Execute " Insert Into NewCaptions Values('05003','Version','05003') "
VstarConn.Execute " Insert Into NewCaptions Values('05004','Print Electronics Equipments Pvt Ltd','05004') "
VstarConn.Execute " Insert Into NewCaptions Values('05005','PEEPL','05005') "
VstarConn.Execute " Insert Into NewCaptions Values('05006','E-mail','05006') "
VstarConn.Execute " Insert Into NewCaptions Values('05007','Support @printelectronics.com','05007') "
VstarConn.Execute " Insert Into NewCaptions Values('05008','sales@printelectronics.com','05008') "
VstarConn.Execute " Insert Into NewCaptions Values('05009','Also Visit us at: www.printelectronics.com','05009') "
VstarConn.Execute " Insert Into NewCaptions Values('05010','Print Electronics Equipment Pvt Ltd','05010') "
VstarConn.Execute " Insert Into NewCaptions Values('06001','Print Electronics Administrator','06001') "
VstarConn.Execute " Insert Into NewCaptions Values('06002','Reset &Exclusive Lock','06002') "
VstarConn.Execute " Insert Into NewCaptions Values('06003','Reset &Daily Lock','06003') "
VstarConn.Execute " Insert Into NewCaptions Values('06004','Reset &Monthly Lock','06004') "
VstarConn.Execute " Insert Into NewCaptions Values('06005','Reset &Yearly Lock','06005') "
VstarConn.Execute " Insert Into NewCaptions Values('06006','Please Confirm that Daily Data is not being Processed elsewhere','06006') "
VstarConn.Execute " Insert Into NewCaptions Values('06007','Please Confirm that Monthly Data is not being Processed elsewhere','06007') "
VstarConn.Execute " Insert Into NewCaptions Values('06008','Please Confirm that Yearly Leaves are not being Processed elsewhere','06008') "
VstarConn.Execute " Insert Into NewCaptions Values('07001','Avail Entry','07001') "
VstarConn.Execute " Insert Into NewCaptions Values('07002','Employee Code','07002') "
VstarConn.Execute " Insert Into NewCaptions Values('07003','Name','07003') "
VstarConn.Execute " Insert Into NewCaptions Values('07004','Leave Information','07004') "
VstarConn.Execute " Insert Into NewCaptions Values('07005','Leave Code','07005') "
VstarConn.Execute " Insert Into NewCaptions Values('07006','Leave type','07006') "
VstarConn.Execute " Insert Into NewCaptions Values('07007','No of Days','07007') "
VstarConn.Execute " Insert Into NewCaptions Values('07008','Balance','07008') "
VstarConn.Execute " Insert Into NewCaptions Values('07009','Leave Code','07009') "
VstarConn.Execute " Insert Into NewCaptions Values('07010','Leave From','07010') "
VstarConn.Execute " Insert Into NewCaptions Values('07011','Leave To','07011') "
VstarConn.Execute " Insert Into NewCaptions Values('07012','Leave Days','07012') "
VstarConn.Execute " Insert Into NewCaptions Values('07013','Code','07013') "
VstarConn.Execute " Insert Into NewCaptions Values('07014','Name','07014') "
VstarConn.Execute " Insert Into NewCaptions Values('07015','Balance','07015') "
VstarConn.Execute " Insert Into NewCaptions Values('07016','Employee Not Found','07016') "
VstarConn.Execute " Insert Into NewCaptions Values('07017','Please Select the Leave to be Availed','07017') "
VstarConn.Execute " Insert Into NewCaptions Values('07018','Leaves cannot be Availed for 0 Number of days','07018') "
VstarConn.Execute " Insert Into NewCaptions Values('07019','Already Availed','07019') "
VstarConn.Execute " Insert Into NewCaptions Values('07020',' Times, Still Continue','07020') "
VstarConn.Execute " Insert Into NewCaptions Values('07021','Maximum','07021') "
VstarConn.Execute " Insert Into NewCaptions Values('07022',' Leave(s) Can be Availed ,Still Continue?','07022') "
VstarConn.Execute " Insert Into NewCaptions Values('07023','Minimum','07023') "
VstarConn.Execute " Insert Into NewCaptions Values('07024','No Leave Balances are Remaining','07024') "
VstarConn.Execute " Insert Into NewCaptions Values('07025','Still Continue ?','07025') "

VstarConn.Execute " Insert Into NewCaptions Values('07026','Balance is already Over, Still Avail Leaves','07026') "
VstarConn.Execute " Insert Into NewCaptions Values('07027','Leaves already Availed on the One of the above Selected Date(s)','07027') "
VstarConn.Execute " Insert Into NewCaptions Values('07028','This Employee is absent on Immediate days.','07028') "
VstarConn.Execute " Insert Into NewCaptions Values('07029',' Leave applied','07029') "
VstarConn.Execute " Insert Into NewCaptions Values('07030','.  Accept this leave?','07030') "
VstarConn.Execute " Insert Into NewCaptions Values('07031','The Employee has Availed Leaves for Continuous Days , Still Continue?','07031') "
VstarConn.Execute " Insert Into NewCaptions Values('07032','Monthly Transaction File not Found','07032') "
VstarConn.Execute " Insert Into NewCaptions Values('07033','Updation Cannot be Done','07033') "
VstarConn.Execute " Insert Into NewCaptions Values('07034',' for Leave Deletion','07034') "
VstarConn.Execute " Insert Into NewCaptions Values('07035','This Leave has been Deleted from the Master','07035') "
VstarConn.Execute " Insert Into NewCaptions Values('07036','This may cause this Form to function improperly','07036') "
VstarConn.Execute " Insert Into NewCaptions Values('07037','Operation Aborted','07037') "
VstarConn.Execute " Insert Into NewCaptions Values('07038','CO for extra work done on','07038') "
VstarConn.Execute " Insert Into NewCaptions Values('07039','Please enter CO Entry Date','07039') "
VstarConn.Execute " Insert Into NewCaptions Values('07040','From and To Date must be same','07040') "
VstarConn.Execute " Insert Into NewCaptions Values('07041','No CO found for specified Date','07041') "
VstarConn.Execute " Insert Into NewCaptions Values('07042','CO Availment Date Limit Over, Still Continue ?','07042') "
VstarConn.Execute " Insert Into NewCaptions Values('07043','CO cannot be availed more than the available balance.','07043') "
VstarConn.Execute " Insert Into NewCaptions Values('07044','Entry Date cannot be as same as From Date','07044') "
VstarConn.Execute " Insert Into NewCaptions Values('08001','BackUp and Restore','08001') "
VstarConn.Execute " Insert Into NewCaptions Values('08002','Operation','08002') "
VstarConn.Execute " Insert Into NewCaptions Values('08003','&BackUp','08003') "
VstarConn.Execute " Insert Into NewCaptions Values('08004','&Restore','08004') "
VstarConn.Execute " Insert Into NewCaptions Values('08005','Please Select the Type of Operation First','08005') "
VstarConn.Execute " Insert Into NewCaptions Values('08006','Please Select the Directory for Back Up','08006') "
VstarConn.Execute " Insert Into NewCaptions Values('08007','Please Select the File to Restore','08007') "
VstarConn.Execute " Insert Into NewCaptions Values('08008','Are You Sure to BackUp the Database','08008') "
VstarConn.Execute " Insert Into NewCaptions Values('08009','Are You Sure to Restore the Database','08009') "
VstarConn.Execute " Insert Into NewCaptions Values('08010','Error Restoring Connection','08010') "
VstarConn.Execute " Insert Into NewCaptions Values('08011','Recommended to Quit and Restart the Application','08011') "
VstarConn.Execute " Insert Into NewCaptions Values('09001','Select Date','09001') "
VstarConn.Execute " Insert Into NewCaptions Values('09002','Please Select the Date First','09002') "
VstarConn.Execute " Insert Into NewCaptions Values('10001','Category Master','10001') "
VstarConn.Execute " Insert Into NewCaptions Values('10002','Category Code','10002') "
VstarConn.Execute " Insert Into NewCaptions Values('10003','Category Name','10003') "
VstarConn.Execute " Insert Into NewCaptions Values('10004','Info','10004') "
VstarConn.Execute " Insert Into NewCaptions Values('10005','Late coming/Early going Rules','10005') "
VstarConn.Execute " Insert Into NewCaptions Values('10006','Allow Employee to come late by','10006') "
VstarConn.Execute " Insert Into NewCaptions Values('10007','Allow Employee to go early by','10007') "
VstarConn.Execute " Insert Into NewCaptions Values('10008','Ignore early arrival before shift by','10008') "
VstarConn.Execute " Insert Into NewCaptions Values('10009','Ignore late going after shift by','10009') "
VstarConn.Execute " Insert Into NewCaptions Values('10010','Cut half day if late coming by','10010') "
VstarConn.Execute " Insert Into NewCaptions Values('10011','Cut half day if early going by','10011') "
VstarConn.Execute " Insert Into NewCaptions Values('10012','Comp .Off Rule for Normal Days','10012') "
VstarConn.Execute " Insert Into NewCaptions Values('10013','Credit half day for working more than','10013') "
VstarConn.Execute " Insert Into NewCaptions Values('10014','Credit full day for working more than','10014') "
VstarConn.Execute " Insert Into NewCaptions Values('10015','Comp .Off Rule for other Days','10015') "
VstarConn.Execute " Insert Into NewCaptions Values('10016','Category Code cannot be blank','10016') "
VstarConn.Execute " Insert Into NewCaptions Values('10017','Category already exists','10017') "
VstarConn.Execute " Insert Into NewCaptions Values('10018','Category with the Same Name Already Exists','10018') "
VstarConn.Execute " Insert Into NewCaptions Values('10019','Category Name cannot be blank','10019') "
VstarConn.Execute " Insert Into NewCaptions Values('11001','Compact Database','11001') "
VstarConn.Execute " Insert Into NewCaptions Values('11002','Note :: This option will Compact the Database file.This will Save the Disk space where the Application file is installed and help in better Application Performance. Before running this option make sure the user has logged in Exclusively and no other use','11002') "
VstarConn.Execute " Insert Into NewCaptions Values('11003','Compact &Database','11003') "
VstarConn.Execute " Insert Into NewCaptions Values('11004','Make Sure you Read the Note on the Form before Running this Option.','11004') "
VstarConn.Execute " Insert Into NewCaptions Values('11005','Database Compacted','11005') "
VstarConn.Execute " Insert Into NewCaptions Values('12001','Company Master','12001') "
VstarConn.Execute " Insert Into NewCaptions Values('12002','Company Code','12002') "
VstarConn.Execute " Insert Into NewCaptions Values('12003','Company Name','12003') "
VstarConn.Execute " Insert Into NewCaptions Values('12004','Company not Found','12004') "
VstarConn.Execute " Insert Into NewCaptions Values('12005','Cannot Add More than','12005') "
VstarConn.Execute " Insert Into NewCaptions Values('12006',' Companies.','12006') "
VstarConn.Execute " Insert Into NewCaptions Values('12007','Company Code cannot be blank','12007') "
VstarConn.Execute " Insert Into NewCaptions Values('12008','Company Code Already Exists','12008') "
VstarConn.Execute " Insert Into NewCaptions Values('12009','Company Name cannot be blank','12009') "
VstarConn.Execute " Insert Into NewCaptions Values('13001','Correction','13001') "
VstarConn.Execute " Insert Into NewCaptions Values('13002','Employee Code','13002') "
VstarConn.Execute " Insert Into NewCaptions Values('13003','Name','13003') "
VstarConn.Execute " Insert Into NewCaptions Values('13004','Attendance Records','13004') "
VstarConn.Execute " Insert Into NewCaptions Values('13005','Attendance Details','13005') "
VstarConn.Execute " Insert Into NewCaptions Values('13006','Work Hrs','13006') "
VstarConn.Execute " Insert Into NewCaptions Values('13007','Present','13007') "
VstarConn.Execute " Insert Into NewCaptions Values('13008','Details','13008') "
VstarConn.Execute " Insert Into NewCaptions Values('13009','Misc.','13009') "
VstarConn.Execute " Insert Into NewCaptions Values('13010','Present Days','13010') "
VstarConn.Execute " Insert Into NewCaptions Values('13011','Rest Hrs','13011') "
VstarConn.Execute " Insert Into NewCaptions Values('13012','CO Days','13012') "
VstarConn.Execute " Insert Into NewCaptions Values('13013','Time','13013') "
VstarConn.Execute " Insert Into NewCaptions Values('13014','Irregular Entries','13014') "
VstarConn.Execute " Insert Into NewCaptions Values('13015','2nd','13015') "
VstarConn.Execute " Insert Into NewCaptions Values('13016','4th','13016') "
VstarConn.Execute " Insert Into NewCaptions Values('13017','6th','13017') "
VstarConn.Execute " Insert Into NewCaptions Values('13018','3rd','13018') "
VstarConn.Execute " Insert Into NewCaptions Values('13019','5th','13019') "
VstarConn.Execute " Insert Into NewCaptions Values('13020','7th','13020') "
VstarConn.Execute " Insert Into NewCaptions Values('13021','On Duty','13021') "
VstarConn.Execute " Insert Into NewCaptions Values('13022','Off Duty','13022') "
VstarConn.Execute " Insert Into NewCaptions Values('13023','Permission','13023') "
VstarConn.Execute " Insert Into NewCaptions Values('13024','Late Card','13024') "
VstarConn.Execute " Insert Into NewCaptions Values('13025','Early Card','13025') "
VstarConn.Execute " Insert Into NewCaptions Values('13026','&Shift','13026') "
VstarConn.Execute " Insert Into NewCaptions Values('13027','&Record','13027') "
VstarConn.Execute " Insert Into NewCaptions Values('13028','&Status','13028') "
VstarConn.Execute " Insert Into NewCaptions Values('13029','&On Duty','13029') "
VstarConn.Execute " Insert Into NewCaptions Values('13030','&Off Duty','13030') "
VstarConn.Execute " Insert Into NewCaptions Values('13031','&Time','13031') "
VstarConn.Execute " Insert Into NewCaptions Values('13032','OT/CO','13032') "
VstarConn.Execute " Insert Into NewCaptions Values('13033','&CO','13033') "
VstarConn.Execute " Insert Into NewCaptions Values('13034','This Employee does not have Overtime or Comp Off','13034') "
VstarConn.Execute " Insert Into NewCaptions Values('13035','File not found for the Month of','13035') "
VstarConn.Execute " Insert Into NewCaptions Values('13036','No Records Found For the Employee','13036') "
VstarConn.Execute " Insert Into NewCaptions Values('13037','Error Finding the Employee Record for the Date','13037') "
VstarConn.Execute " Insert Into NewCaptions Values('13038','Invalid Value','13038') "
VstarConn.Execute " Insert Into NewCaptions Values('13039','Please Select the Shift','13039') "
VstarConn.Execute " Insert Into NewCaptions Values('13040','Invalid Late Card Number','13040') "
VstarConn.Execute " Insert Into NewCaptions Values('13041','Invalid Early Card Number','13041') "
VstarConn.Execute " Insert Into NewCaptions Values('13042','Minutes Should be less than 60','13042') "
VstarConn.Execute " Insert Into NewCaptions Values('13043','Invalid value :: Cannot be Greater than 48','13043') "
VstarConn.Execute " Insert Into NewCaptions Values('13044','Departure Time Cannot be 0 if Arrival Time is Greater than 0','13044') "
VstarConn.Execute " Insert Into NewCaptions Values('13045','Arrival Time Cannot be Greater then Departure Time','13045') "
VstarConn.Execute " Insert Into NewCaptions Values('13046','Arrival Time Cannot be 0 if Departure Time is Greater than 0','13046') "
VstarConn.Execute " Insert Into NewCaptions Values('13047','Punch Time Should be between Arrival Time and Departure Time','13047') "
VstarConn.Execute " Insert Into NewCaptions Values('13048','To Time Should be Greater than From Time','13048') "
VstarConn.Execute " Insert Into NewCaptions Values('13049','On Duty From Time Cannot Be Greater than On Duty To Time','13049') "
VstarConn.Execute " Insert Into NewCaptions Values('13050','On Duty From punch Missing','13050') "
VstarConn.Execute " Insert Into NewCaptions Values('13051','On Duty From Time Should be between Arrival Time and Departure Time','13051') "
VstarConn.Execute " Insert Into NewCaptions Values('13052','On Duty To Time Should be between Arrival Time and Departure Time','13052') "
VstarConn.Execute " Insert Into NewCaptions Values('13053','Off Duty From punch Missing','13053') "
VstarConn.Execute " Insert Into NewCaptions Values('13054','Off Duty From Time Should be between Arrival Time and Departure Time','13054') "
VstarConn.Execute " Insert Into NewCaptions Values('13055','Off Duty To Time Should be between Arrival Time and Departure Time','13055') "
VstarConn.Execute " Insert Into NewCaptions Values('13056','2nd punch Missing','13056') "
VstarConn.Execute " Insert Into NewCaptions Values('13057','3rd punch Missing','13057') "
VstarConn.Execute " Insert Into NewCaptions Values('13058','4th punch Missing','13058') "
VstarConn.Execute " Insert Into NewCaptions Values('13059','5th punch Missing','13059') "
VstarConn.Execute " Insert Into NewCaptions Values('13060','5th punch Missing','13060') "
VstarConn.Execute " Insert Into NewCaptions Values('13061','2nd punch cannot be Greater than 3rd Punch','13061') "
VstarConn.Execute " Insert Into NewCaptions Values('13062','3rd punch cannot be Greater than 4th Punch','13062') "
VstarConn.Execute " Insert Into NewCaptions Values('13063','4th punch cannot be Greater than 5th Punch','13063') "
VstarConn.Execute " Insert Into NewCaptions Values('13064','5th punch cannot be Greater than 6th Punch','13064') "
VstarConn.Execute " Insert Into NewCaptions Values('13065','6th punch cannot be Greater than 7th Punch','13065') "
VstarConn.Execute " Insert Into NewCaptions Values('13066','CO not Found :: Leave Balance File for the Current Year not Updated','13066') "
VstarConn.Execute " Insert Into NewCaptions Values('13067','CO not Found in Leave Master','13067') "
VstarConn.Execute " Insert Into NewCaptions Values('13068','Leave Balanace File for the Current Year not Found','13068') "
VstarConn.Execute " Insert Into NewCaptions Values('13069','Please Create it First and then do the Daily Process','13069') "
VstarConn.Execute " Insert Into NewCaptions Values('13070','NO CO Rule is set','13070') "
VstarConn.Execute " Insert Into NewCaptions Values('14001','Change Period','14001') "
VstarConn.Execute " Insert Into NewCaptions Values('14002','From Day','14002') "
VstarConn.Execute " Insert Into NewCaptions Values('14003','To Day','14003') "
VstarConn.Execute " Insert Into NewCaptions Values('14004','&Overwtite Week Off''s','14004') "
VstarConn.Execute " Insert Into NewCaptions Values('14005','Overwrite &Holidays','14005') "
VstarConn.Execute " Insert Into NewCaptions Values('14006','&Change','14006') "
VstarConn.Execute " Insert Into NewCaptions Values('14007','Periodic Shift Updation done','14007') "
VstarConn.Execute " Insert Into NewCaptions Values('14008','Please Select the Month First','14008') "
VstarConn.Execute " Insert Into NewCaptions Values('14009','Please Select the Year First','14009') "
VstarConn.Execute " Insert Into NewCaptions Values('14010','Shift File not found for the Month of','14010') "
VstarConn.Execute " Insert Into NewCaptions Values('14011','Please Create it First Using Shift Creation','14011') "
VstarConn.Execute " Insert Into NewCaptions Values('14012','Please Select the Day from the where Shifts are to be Updated','14012') "
VstarConn.Execute " Insert Into NewCaptions Values('14013','Please Select the Day Till the where Shifts are to be Updated','14013') "
VstarConn.Execute " Insert Into NewCaptions Values('14014','From Period cannot be Greater than To Period','14014') "
VstarConn.Execute " Insert Into NewCaptions Values('14015','Please Select the Shift','14015') "
VstarConn.Execute " Insert Into NewCaptions Values('14016','Please Select the Employee First','14016') "
VstarConn.Execute " Insert Into NewCaptions Values('15001','Credit Entry','15001') "
VstarConn.Execute " Insert Into NewCaptions Values('15002','Credit On','15002') "
VstarConn.Execute " Insert Into NewCaptions Values('15003','Please Select the Leave to be Credited','15003') "
VstarConn.Execute " Insert Into NewCaptions Values('15004','Leaves cannot be Credited for 0 Number of days','15004') "
VstarConn.Execute " Insert Into NewCaptions Values('15005','Days To be Credited Must be Divisible by 0.50','15005') "
VstarConn.Execute " Insert Into NewCaptions Values('15006','From date Cannot be Empty','15006') "
VstarConn.Execute " Insert Into NewCaptions Values('15007','Leaves already Credited on the above Selected Date','15007') "
VstarConn.Execute " Insert Into NewCaptions Values('15008','Maximum Credit every year are','15008') "
VstarConn.Execute " Insert Into NewCaptions Values('15009',' days','15009') "
VstarConn.Execute " Insert Into NewCaptions Values('15010','Credit All days?','15010') "
VstarConn.Execute " Insert Into NewCaptions Values('16001','Customize leave Codes','16001') "
VstarConn.Execute " Insert Into NewCaptions Values('16002','Keep default Codes for present /Absent/Week Off /Holiday','16002') "
VstarConn.Execute " Insert Into NewCaptions Values('16003','Change Leave Codes for present /Absent/Week Off /Holiday','16003') "
VstarConn.Execute " Insert Into NewCaptions Values('16004','Type','16004') "
VstarConn.Execute " Insert Into NewCaptions Values('16005','Existing Codes','16005') "
VstarConn.Execute " Insert Into NewCaptions Values('16006','New Codes','16006') "
VstarConn.Execute " Insert Into NewCaptions Values('16007','Absent Days','16007') "
VstarConn.Execute " Insert Into NewCaptions Values('16008','Present Days','16008') "
VstarConn.Execute " Insert Into NewCaptions Values('16009','Week Offs','16009') "
VstarConn.Execute " Insert Into NewCaptions Values('16010','Holidays','16010') "
VstarConn.Execute " Insert Into NewCaptions Values('16011','&Save and Exit','16011') "
VstarConn.Execute " Insert Into NewCaptions Values('16012','Please Enter the Absent Code','16012') "
VstarConn.Execute " Insert Into NewCaptions Values('16013','Please Enter the Present Code','16013') "
VstarConn.Execute " Insert Into NewCaptions Values('16014','Please Enter the Week Off Code','16014') "
VstarConn.Execute " Insert Into NewCaptions Values('16015','Please Enter the Holidays Code','16015') "
VstarConn.Execute " Insert Into NewCaptions Values('16016','Duplicate codes not Allowed','16016') "
VstarConn.Execute " Insert Into NewCaptions Values('16017','Are You Sure to Change the Custom Codes','16017') "
VstarConn.Execute " Insert Into NewCaptions Values('16018','Error in Updating Custom Codes','16018') "
VstarConn.Execute " Insert Into NewCaptions Values('16019','Please Create Yearly Leave Files','16019') "
VstarConn.Execute " Insert Into NewCaptions Values('17001','Daily Processing','17001') "
VstarConn.Execute " Insert Into NewCaptions Values('17002','Processing Dates','17002') "
VstarConn.Execute " Insert Into NewCaptions Values('17003','&From Date','17003') "
VstarConn.Execute " Insert Into NewCaptions Values('17004','&To Date','17004') "
VstarConn.Execute " Insert Into NewCaptions Values('17005','Select Dat File','17005') "
VstarConn.Execute " Insert Into NewCaptions Values('17006','&Exclude Dat Files','17006') "
VstarConn.Execute " Insert Into NewCaptions Values('17007','Retreiving Records from the Dat File ..','17007') "
VstarConn.Execute " Insert Into NewCaptions Values('17008','Processing Records :: Please Wait ..','17008') "
VstarConn.Execute " Insert Into NewCaptions Values('17009','&Process','17009') "
VstarConn.Execute " Insert Into NewCaptions Values('17010','This Will Clear your All Dat Files Selection','17010') "
VstarConn.Execute " Insert Into NewCaptions Values('17011','Daily Process is Aborted','17011') "
VstarConn.Execute " Insert Into NewCaptions Values('17012','Daliy Process is Over','17012') "
VstarConn.Execute " Insert Into NewCaptions Values('17013','Software Locked :: Cannot Process','17013') "
VstarConn.Execute " Insert Into NewCaptions Values('17014','Contact Print Electronics','17014') "
VstarConn.Execute " Insert Into NewCaptions Values('17015','Duplicate File Names not Allowed','17015') "
VstarConn.Execute " Insert Into NewCaptions Values('17016','Since all the Shift Files Necessary for Processing are not Created ,Processing Cannot Continue.','17016') "
VstarConn.Execute " Insert Into NewCaptions Values('17017','Shift file for the month of','17017') "
VstarConn.Execute " Insert Into NewCaptions Values('17018',' not available','17018') "
VstarConn.Execute " Insert Into NewCaptions Values('17019','Do you want to create it','17019') "
VstarConn.Execute " Insert Into NewCaptions Values('17020','Please do the Processing for','17020') "

VstarConn.Execute " Insert Into NewCaptions Values('17021','Remove','17021') "
VstarConn.Execute " Insert Into NewCaptions Values('17022','Click to Toggle Selection','17022') "
VstarConn.Execute " Insert Into NewCaptions Values('18001','Select Source','18001') "
VstarConn.Execute " Insert Into NewCaptions Values('18002','&Files','18002') "
VstarConn.Execute " Insert Into NewCaptions Values('18003','&Drives','18003') "
VstarConn.Execute " Insert Into NewCaptions Values('18004','Fold&ers','18004') "
VstarConn.Execute " Insert Into NewCaptions Values('19001','Declare Holiday/ WeekOff','19001') "
VstarConn.Execute " Insert Into NewCaptions Values('19002','Compensate On','19002') "
VstarConn.Execute " Insert Into NewCaptions Values('19003','As','19003') "
VstarConn.Execute " Insert Into NewCaptions Values('19004','Add this Holiday/WeekOff for all categories','19004') "
VstarConn.Execute " Insert Into NewCaptions Values('19005','Compensate date','19005') "
VstarConn.Execute " Insert Into NewCaptions Values('19006','Declare as','19006') "
VstarConn.Execute " Insert Into NewCaptions Values('19007','WeekOff','19007') "
VstarConn.Execute " Insert Into NewCaptions Values('19008','Holiday','19008') "
VstarConn.Execute " Insert Into NewCaptions Values('19009','Category Does not Exist :: Cannot Display the Record','19009') "
VstarConn.Execute " Insert Into NewCaptions Values('19010','Category cannot be blank','19010') "
VstarConn.Execute " Insert Into NewCaptions Values('19011','Blank Category Master  :: Cannot Add the Record','19011') "
VstarConn.Execute " Insert Into NewCaptions Values('19012','Date cannot be blank','19012') "
VstarConn.Execute " Insert Into NewCaptions Values('19013','Holiday/Week Off Date','19013') "
VstarConn.Execute " Insert Into NewCaptions Values('19014','Compensate Date','19014') "
VstarConn.Execute " Insert Into NewCaptions Values('19015','Holiday Date and Compensate Date cannot be Same','19015') "
VstarConn.Execute " Insert Into NewCaptions Values('19016','Description cannot be Blank','19016') "
VstarConn.Execute " Insert Into NewCaptions Values('19017','It''s a Week Off,Cannot Declare Holiday/Week Off on the Same Date','19017') "
VstarConn.Execute " Insert Into NewCaptions Values('19018','Holiday Already Declared on the Selected Date','19018') "
VstarConn.Execute " Insert Into NewCaptions Values('19019','No Employees :: Cannot Add Holidays.','19019') "
VstarConn.Execute " Insert Into NewCaptions Values('20001','DEPARTMENT MASTER','20001') "
VstarConn.Execute " Insert Into NewCaptions Values('20002','Strength','20002') "
VstarConn.Execute " Insert Into NewCaptions Values('20003','Department not Found','20003') "
VstarConn.Execute " Insert Into NewCaptions Values('20004','Department Code cannot be blank','20004') "
VstarConn.Execute " Insert Into NewCaptions Values('20005','Department Code Already Exists','20005') "
VstarConn.Execute " Insert Into NewCaptions Values('20006','Department Name cannot be blank','20006') "
VstarConn.Execute " Insert Into NewCaptions Values('20007','Department with Same Code Already Exists','20007') "
VstarConn.Execute " Insert Into NewCaptions Values('21001','Make DSN','21001') "
VstarConn.Execute " Insert Into NewCaptions Values('21002','&Back End','21002') "
VstarConn.Execute " Insert Into NewCaptions Values('21003','&User Name','21003') "
VstarConn.Execute " Insert Into NewCaptions Values('21004','&Password','21004') "
VstarConn.Execute " Insert Into NewCaptions Values('21005','DSN Name','21005') "
VstarConn.Execute " Insert Into NewCaptions Values('21006','Server &Name','21006') "
VstarConn.Execute " Insert Into NewCaptions Values('21007','P&ath','21007') "
VstarConn.Execute " Insert Into NewCaptions Values('21008','Due to Some Reasons the DSN may be Corrupted or Deleted. DSN can be Created with a Valid User Name,Password (Case Sensitive) and a Server Name. Please Contact Your System Administrator or Print Electronics for further Details.','21008') "
VstarConn.Execute " Insert Into NewCaptions Values('21009','Due to Some Reasons the DSN may be Corrupted or Deleted. DSN can be Created with a Valid Password(Case Sensitive) ,Also Enter a Valid MDB File Path. Please Contact Your System Administrator or Print Electronics for further Details.','21009') "
VstarConn.Execute " Insert Into NewCaptions Values('21010','Details not Yet Available Yet. Please Contact Print Electronics for further Details.','21010') "
VstarConn.Execute " Insert Into NewCaptions Values('21011','&Show System DSN Wizard','21011') "
VstarConn.Execute " Insert Into NewCaptions Values('21012','Please Enter Server Name','21012') "
VstarConn.Execute " Insert Into NewCaptions Values('21013','Please Enter UserName','21013') "
VstarConn.Execute " Insert Into NewCaptions Values('21014','Please Enter Password','21014') "
VstarConn.Execute " Insert Into NewCaptions Values('21015','Please Enter Database Path','21015') "
VstarConn.Execute " Insert Into NewCaptions Values('21016','DSN Created Successfully','21016') "
VstarConn.Execute " Insert Into NewCaptions Values('22001','Edit Paid Days','22001') "
VstarConn.Execute " Insert Into NewCaptions Values('22002','Employee Code','22002') "
VstarConn.Execute " Insert Into NewCaptions Values('22003','Paid Days','22003') "
VstarConn.Execute " Insert Into NewCaptions Values('22004','Present','22004') "
VstarConn.Execute " Insert Into NewCaptions Values('22005','Absent','22005') "
VstarConn.Execute " Insert Into NewCaptions Values('22006','WeekOff','22006') "
VstarConn.Execute " Insert Into NewCaptions Values('22007','Holiday','22007') "
VstarConn.Execute " Insert Into NewCaptions Values('22008','Please Create it First','22008') "
VstarConn.Execute " Insert Into NewCaptions Values('23001','Employee Master','23001') "
VstarConn.Execute " Insert Into NewCaptions Values('23002','Find Employee with Employee code','23002') "
VstarConn.Execute " Insert Into NewCaptions Values('23003','or having name','23003') "
VstarConn.Execute " Insert Into NewCaptions Values('23004','Official Details','23004') "
VstarConn.Execute " Insert Into NewCaptions Values('23005','Personal Details','23005') "
VstarConn.Execute " Insert Into NewCaptions Values('23006','Other Details','23006') "
VstarConn.Execute " Insert Into NewCaptions Values('23007','Emp Code','23007') "
VstarConn.Execute " Insert Into NewCaptions Values('23008','Employee Name','23008') "
VstarConn.Execute " Insert Into NewCaptions Values('23009','Emp Card','23009') "
VstarConn.Execute " Insert Into NewCaptions Values('23010','Join Date','23010') "
VstarConn.Execute " Insert Into NewCaptions Values('23011','Conf. Date','23011') "
VstarConn.Execute " Insert Into NewCaptions Values('23012','Code No','23012') "
VstarConn.Execute " Insert Into NewCaptions Values('23013','Card No','23013') "
VstarConn.Execute " Insert Into NewCaptions Values('23014','Designation','23014') "
VstarConn.Execute " Insert Into NewCaptions Values('23015','Identification','23015') "
VstarConn.Execute " Insert Into NewCaptions Values('23016','Min. Entry','23016') "
VstarConn.Execute " Insert Into NewCaptions Values('23017','Compensatory Off','23017') "
VstarConn.Execute " Insert Into NewCaptions Values('23018','Autoshift Change','23018') "
VstarConn.Execute " Insert Into NewCaptions Values('23019','Travel By','23019') "
VstarConn.Execute " Insert Into NewCaptions Values('23020','Division','23020') "
VstarConn.Execute " Insert Into NewCaptions Values('23021','Working Schedule','23021') "
VstarConn.Execute " Insert Into NewCaptions Values('23022','Define Schedule','23022') "
VstarConn.Execute " Insert Into NewCaptions Values('23023','Past Employee','23023') "
VstarConn.Execute " Insert Into NewCaptions Values('23024','Left Date','23024') "
VstarConn.Execute " Insert Into NewCaptions Values('23025','Date of Birth','23025') "
VstarConn.Execute " Insert Into NewCaptions Values('23026','Blood Group','23026') "
VstarConn.Execute " Insert Into NewCaptions Values('23027','Date of join','23027') "
VstarConn.Execute " Insert Into NewCaptions Values('23028','Confirm Date','23028') "
VstarConn.Execute " Insert Into NewCaptions Values('23029','Sex','23029') "
VstarConn.Execute " Insert Into NewCaptions Values('23030','E-Mail ID','23030') "
VstarConn.Execute " Insert Into NewCaptions Values('23031','Basic Salary','23031') "
VstarConn.Execute " Insert Into NewCaptions Values('23032','Reference','23032') "
VstarConn.Execute " Insert Into NewCaptions Values('23033','Address','23033') "
VstarConn.Execute " Insert Into NewCaptions Values('23034','City','23034') "
VstarConn.Execute " Insert Into NewCaptions Values('23035','Pin Code','23035') "
VstarConn.Execute " Insert Into NewCaptions Values('23036','Phone No','23036') "
VstarConn.Execute " Insert Into NewCaptions Values('23037','Permanent Address','23037') "
VstarConn.Execute " Insert Into NewCaptions Values('23038','HouseNo/Name','23038') "
VstarConn.Execute " Insert Into NewCaptions Values('23039','City/Village','23039') "
VstarConn.Execute " Insert Into NewCaptions Values('23040','District','23040') "
VstarConn.Execute " Insert Into NewCaptions Values('23041','Tel.No','23041') "
VstarConn.Execute " Insert Into NewCaptions Values('23042','Area','23042') "
VstarConn.Execute " Insert Into NewCaptions Values('23043','Road','23043') "
VstarConn.Execute " Insert Into NewCaptions Values('23044','State','23044') "
VstarConn.Execute " Insert Into NewCaptions Values('23045','Nationality','23045') "
VstarConn.Execute " Insert Into NewCaptions Values('23046','Special Comments','23046') "
VstarConn.Execute " Insert Into NewCaptions Values('23047','Record not Found','23047') "
VstarConn.Execute " Insert Into NewCaptions Values('23048','Please Enter Employee Code','23048') "
VstarConn.Execute " Insert Into NewCaptions Values('23049','Yearly Leave Files are Not Created :: Please Create Them','23049') "
VstarConn.Execute " Insert Into NewCaptions Values('23050','Maximum Employee(s) Allowed :','23050') "
VstarConn.Execute " Insert Into NewCaptions Values('23051','Employee Already Exists','23051') "
VstarConn.Execute " Insert Into NewCaptions Values('23052','Employee Card Number Should be of','23052') "
VstarConn.Execute " Insert Into NewCaptions Values('23053',' Characters','23053') "
VstarConn.Execute " Insert Into NewCaptions Values('23054','Card Number Already Exists','23054') "
VstarConn.Execute " Insert Into NewCaptions Values('23055','Please Enter Employee Name','23055') "
VstarConn.Execute " Insert Into NewCaptions Values('23056','Please Select Employee Category','23056') "
VstarConn.Execute " Insert Into NewCaptions Values('23057','Please Select Employee Department','23057') "
VstarConn.Execute " Insert Into NewCaptions Values('23058','Please Select Employee Group','23058') "
VstarConn.Execute " Insert Into NewCaptions Values('23059','Please Select Company Code','23059') "
VstarConn.Execute " Insert Into NewCaptions Values('23060','Please Define Employee Shift','23060') "
VstarConn.Execute " Insert Into NewCaptions Values('23061','Please Enter Employee Joindate','23061') "
VstarConn.Execute " Insert Into NewCaptions Values('23062','Employee Joindate Must be Less than Employee Shift Date','23062') "
VstarConn.Execute " Insert Into NewCaptions Values('23063','Employee Joindate Must be Less than Employee Leave Date','23063') "
VstarConn.Execute " Insert Into NewCaptions Values('23064','Employee Shift Date Must be Less than Employee Leave Date','23064') "
VstarConn.Execute " Insert Into NewCaptions Values('23065','Employee Birth Date Must be Less than Employee Join Date','23065') "
VstarConn.Execute " Insert Into NewCaptions Values('23066','Employee Confirm Date Must be Greater than Employee Join Date','23066') "
VstarConn.Execute " Insert Into NewCaptions Values('23067',' is a reserved Permission Card No.','23067') "
VstarConn.Execute " Insert Into NewCaptions Values('23068','Please Select OT Rule','23068') "
VstarConn.Execute " Insert Into NewCaptions Values('23069','Please Select CO Rule','23069') "
VstarConn.Execute " Insert Into NewCaptions Values('23070','Please Select Location Code','23070') "
VstarConn.Execute " Insert Into NewCaptions Values('24001','Schedule Master','24001') "
VstarConn.Execute " Insert Into NewCaptions Values('24002','General','24002') "
VstarConn.Execute " Insert Into NewCaptions Values('24003','Shift Info','24003') "
VstarConn.Execute " Insert Into NewCaptions Values('24004','Shift Type','24004') "
VstarConn.Execute " Insert Into NewCaptions Values('24005','Starting Date','24005') "
VstarConn.Execute " Insert Into NewCaptions Values('24006','Rotation Code','24006') "
VstarConn.Execute " Insert Into NewCaptions Values('24007','Shift Code','24007') "
VstarConn.Execute " Insert Into NewCaptions Values('24008','WeekOff','24008') "
VstarConn.Execute " Insert Into NewCaptions Values('24009','There is a weekOff on every','24009') "
VstarConn.Execute " Insert Into NewCaptions Values('24010','Of a week','24010') "
VstarConn.Execute " Insert Into NewCaptions Values('24011','Additional Week-Offs','24011') "
VstarConn.Execute " Insert Into NewCaptions Values('24012','There is a week Off every','24012') "
VstarConn.Execute " Insert Into NewCaptions Values('24013','There is a week Off on the first && third','24013') "
VstarConn.Execute " Insert Into NewCaptions Values('24014','There is a week Off on the second && fourth','24014') "
VstarConn.Execute " Insert Into NewCaptions Values('24015','Shift Date Cannot be Empty','24015') "
VstarConn.Execute " Insert Into NewCaptions Values('24016','ShifDate Cannot be Less then the Join date','24016') "
VstarConn.Execute " Insert Into NewCaptions Values('24017','Please Select the Type of Shift','24017') "
VstarConn.Execute " Insert Into NewCaptions Values('24018','Please Select the Shift','24018') "
VstarConn.Execute " Insert Into NewCaptions Values('24019','Please Select the Rotational Shift','24019') "
VstarConn.Execute " Insert Into NewCaptions Values('24020','Please Select the Week Off Before Selecting the Additional week Off','24020') "
VstarConn.Execute " Insert Into NewCaptions Values('24021','Please Select the Additional Week Off','24021') "
VstarConn.Execute " Insert Into NewCaptions Values('24022','Please Select the First and Third Week Off','24022') "
VstarConn.Execute " Insert Into NewCaptions Values('24023','Please Select the Second and Fourth Week Off','24023') "
VstarConn.Execute " Insert Into NewCaptions Values('24024','Details regarding Daily Processing','24024') "
VstarConn.Execute " Insert Into NewCaptions Values('24025','On Weekoff / Holiday do the following','24025') "
VstarConn.Execute " Insert Into NewCaptions Values('24026','Assign Previous day Shift (Transaction)','24026') "
VstarConn.Execute " Insert Into NewCaptions Values('24027','Assign Next day Shift (Schedule)','24027') "
VstarConn.Execute " Insert Into NewCaptions Values('24028','Assign the following Shift','24028') "
VstarConn.Execute " Insert Into NewCaptions Values('24029','Assign Auto shift if punch found','24029') "
VstarConn.Execute " Insert Into NewCaptions Values('24030','If Blank Shift found','24030') "
VstarConn.Execute " Insert Into NewCaptions Values('24031','Keep it blank','24031') "
VstarConn.Execute " Insert Into NewCaptions Values('24032','Assign this Shift','24032') "
VstarConn.Execute " Insert Into NewCaptions Values('24033','Please Select the Shift to be assigned for Week Off / Holiday','24033') "
VstarConn.Execute " Insert Into NewCaptions Values('24034','Please Select the Shift to be assigned if  Blank Shift is Found','24034') "
VstarConn.Execute " Insert Into NewCaptions Values('24035','    &Set for more employees','24035') "
VstarConn.Execute " Insert Into NewCaptions Values('25001','Encash Entry','25001') "
VstarConn.Execute " Insert Into NewCaptions Values('25002','Encash on','25002') "
VstarConn.Execute " Insert Into NewCaptions Values('25003','Please Select the Leave to be Encashed','25003') "
VstarConn.Execute " Insert Into NewCaptions Values('25004','Leaves cannot be Encashed for 0 Number of days','25004') "
VstarConn.Execute " Insert Into NewCaptions Values('25005','Days To be Encashed Must be Divisible by 0.50','25005') "
VstarConn.Execute " Insert Into NewCaptions Values('25006','Balance is already Over, Still Encash Leaves','25006') "
VstarConn.Execute " Insert Into NewCaptions Values('25007','Leaves already Encashed on the above Selected Date','25007') "
VstarConn.Execute " Insert Into NewCaptions Values('26001','Select Dat File','26001') "
VstarConn.Execute " Insert Into NewCaptions Values('26002','Remove','26002') "
VstarConn.Execute " Insert Into NewCaptions Values('26003','Daily Processing Done for','26003') "
VstarConn.Execute " Insert Into NewCaptions Values('26004','No Valid Entries Found in the Dat File(s) for','26004') "
VstarConn.Execute " Insert Into NewCaptions Values('28001','Administrative User','28001') "
VstarConn.Execute " Insert Into NewCaptions Values('28002','This User Will Have All the Administrative Rights','28002') "
VstarConn.Execute " Insert Into NewCaptions Values('28003','&User Name','28003') "
VstarConn.Execute " Insert Into NewCaptions Values('28004','&Password','28004') "
VstarConn.Execute " Insert Into NewCaptions Values('28005','&Re-Type Password','28005') "
VstarConn.Execute " Insert Into NewCaptions Values('28006','The Date Format Selected will Effect the Software throughout it''s Life Time.','28006') "
VstarConn.Execute " Insert Into NewCaptions Values('28007','&British e.g. 29/03/01','28007') "
VstarConn.Execute " Insert Into NewCaptions Values('28008','&American e.g. 03/29/01','28008') "
VstarConn.Execute " Insert Into NewCaptions Values('28009','Date Format','28009') "
VstarConn.Execute " Insert Into NewCaptions Values('28010','The Date Format You Have Selected is British','28010') "
VstarConn.Execute " Insert Into NewCaptions Values('28011','Are you Sure to Continue ?','28011') "
VstarConn.Execute " Insert Into NewCaptions Values('28012','The Date Format You Have Selected is American','28012') "
VstarConn.Execute " Insert Into NewCaptions Values('28013','The Passwords don''t Match Please Re-Enter','28013') "
VstarConn.Execute " Insert Into NewCaptions Values('28014','Encryption Error :: Try Another Password','28014') "
VstarConn.Execute " Insert Into NewCaptions Values('28015',' Added as Admin Successfully','28015') "
VstarConn.Execute " Insert Into NewCaptions Values('29001','Group Master','29001') "
VstarConn.Execute " Insert Into NewCaptions Values('29002','Group Code cannot be Blank','29002') "
VstarConn.Execute " Insert Into NewCaptions Values('29003','Group Already Exists','29003') "
VstarConn.Execute " Insert Into NewCaptions Values('29004','Group Description cannot be Blank','29004') "
VstarConn.Execute " Insert Into NewCaptions Values('30001','Holiday Master','30001') "
VstarConn.Execute " Insert Into NewCaptions Values('30002','Holiday Date','30002') "
VstarConn.Execute " Insert Into NewCaptions Values('30003','Name of Holiday','30003') "
VstarConn.Execute " Insert Into NewCaptions Values('30004','Add this holiday for all categories','30004') "
VstarConn.Execute " Insert Into NewCaptions Values('30005','Specific Category','30005') "
VstarConn.Execute " Insert Into NewCaptions Values('30006','Category cannot be blank','30006') "
VstarConn.Execute " Insert Into NewCaptions Values('30007','Blank Category Master  :: Cannot Add the Record','30007') "
VstarConn.Execute " Insert Into NewCaptions Values('30008','Date cannot be blank','30008') "
VstarConn.Execute " Insert Into NewCaptions Values('30009','Description cannot be Blank','30009') "

VstarConn.Execute " Insert Into NewCaptions Values('30010','Holiday Already mentioned for this category','30010') "
VstarConn.Execute " Insert Into NewCaptions Values('30011','Category Does not Exist :: Cannot Display the Record','30011') "
VstarConn.Execute " Insert Into NewCaptions Values('31001','Leave Master','31001') "
VstarConn.Execute " Insert Into NewCaptions Values('31002','Custom Codes','31002') "
VstarConn.Execute " Insert Into NewCaptions Values('31003','Leave Code','31003') "
VstarConn.Execute " Insert Into NewCaptions Values('31004','Name of Leave','31004') "
VstarConn.Execute " Insert Into NewCaptions Values('31005','Leave Balance','31005') "
VstarConn.Execute " Insert Into NewCaptions Values('31006','Definition','31006') "
VstarConn.Execute " Insert Into NewCaptions Values('31007','Count this leave in payable days','31007') "
VstarConn.Execute " Insert Into NewCaptions Values('31008','Keep balance for this leave','31008') "
VstarConn.Execute " Insert Into NewCaptions Values('31009','For the current year','31009') "
VstarConn.Execute " Insert Into NewCaptions Values('31010','Credit','31010') "
VstarConn.Execute " Insert Into NewCaptions Values('31011','days leaves for consumption','31011') "
VstarConn.Execute " Insert Into NewCaptions Values('31012','Allow maximum','31012') "
VstarConn.Execute " Insert Into NewCaptions Values('31013','days leave to be accumulated','31013') "
VstarConn.Execute " Insert Into NewCaptions Values('31014','Crediting for new employees','31014') "
VstarConn.Execute " Insert Into NewCaptions Values('31015','Credit immediately','31015') "
VstarConn.Execute " Insert Into NewCaptions Values('31016','Credit next year','31016') "
VstarConn.Execute " Insert Into NewCaptions Values('31017','While crediting','31017') "
VstarConn.Execute " Insert Into NewCaptions Values('31018','Credit leaves full','31018') "
VstarConn.Execute " Insert Into NewCaptions Values('31019','Credit in proportion','31019') "
VstarConn.Execute " Insert Into NewCaptions Values('31020','All Categories','31020') "
VstarConn.Execute " Insert Into NewCaptions Values('31021','Specific','31021') "
VstarConn.Execute " Insert Into NewCaptions Values('31022','Mark Leaves','31022') "
VstarConn.Execute " Insert Into NewCaptions Values('31023','Including weekOff/Holidays','31023') "
VstarConn.Execute " Insert Into NewCaptions Values('31024','Excluding weekOff/Holidays','31024') "
VstarConn.Execute " Insert Into NewCaptions Values('31025','Decide while entering leave','31025') "
VstarConn.Execute " Insert Into NewCaptions Values('31026','At the End of the Year','31026') "
VstarConn.Execute " Insert Into NewCaptions Values('31027','Carry forward balance leaves','31027') "
VstarConn.Execute " Insert Into NewCaptions Values('31028','Encash balance leaves','31028') "
VstarConn.Execute " Insert Into NewCaptions Values('31029','Check following rules','31029') "
VstarConn.Execute " Insert Into NewCaptions Values('31030','Allow','31030') "
VstarConn.Execute " Insert Into NewCaptions Values('31031','times in ayear','31031') "
VstarConn.Execute " Insert Into NewCaptions Values('31032','Maximum','31032') "
VstarConn.Execute " Insert Into NewCaptions Values('31033','days at a time','31033') "
VstarConn.Execute " Insert Into NewCaptions Values('31034','Minimum','31034') "
VstarConn.Execute " Insert Into NewCaptions Values('31035','This Leave codes Defined Once will effect the Entire Software.','31035') "
VstarConn.Execute " Insert Into NewCaptions Values('31036','You Cannot Change them again or cannot make them Default.','31036') "
VstarConn.Execute " Insert Into NewCaptions Values('31037','Proceed further Y/N ?','31037') "
VstarConn.Execute " Insert Into NewCaptions Values('31038','Leave Code should be of atleast 2 Characters','31038') "
VstarConn.Execute " Insert Into NewCaptions Values('31039','Leave Code must not be Equal to ABSENT,PRESENT,WEEK OFF or HOLIDAY Code','31039') "
VstarConn.Execute " Insert Into NewCaptions Values('31040','Please Add Categories First','31040') "
VstarConn.Execute " Insert Into NewCaptions Values('31041','Please Select the Category First','31041') "
VstarConn.Execute " Insert Into NewCaptions Values('31042','Leave Already Defined','31042') "
VstarConn.Execute " Insert Into NewCaptions Values('31043','Leave Name cannot be Blank','31043') "
VstarConn.Execute " Insert Into NewCaptions Values('31044','Leave Acculumation should be Greater than Leave Credited','31044') "
VstarConn.Execute " Insert Into NewCaptions Values('31045','Please Select Full or Proportionate Leave','31045') "
VstarConn.Execute " Insert Into NewCaptions Values('31046','This may effect other Leave Specific Details throughout the Application','31046') "
VstarConn.Execute " Insert Into NewCaptions Values('32001','Lost Entry','32001') "
VstarConn.Execute " Insert Into NewCaptions Values('32002','Date of Punch','32002') "
VstarConn.Execute " Insert Into NewCaptions Values('32003','Time of Punch','32003') "
VstarConn.Execute " Insert Into NewCaptions Values('32004','Employee Name','32004') "
VstarConn.Execute " Insert Into NewCaptions Values('32005','Employee Cannot be blank','32005') "
VstarConn.Execute " Insert Into NewCaptions Values('32006','Time can''t be zero','32006') "
VstarConn.Execute " Insert Into NewCaptions Values('33001','Yearly Leave File Updation','33001') "
VstarConn.Execute " Insert Into NewCaptions Values('33002','Instructions','33002') "
VstarConn.Execute " Insert Into NewCaptions Values('33003','Yearly updation transfer leave balances for next year','33003') "
VstarConn.Execute " Insert Into NewCaptions Values('33004','Use only once at the beginning of the year','33004') "
VstarConn.Execute " Insert Into NewCaptions Values('33005','Leave Code','33005') "
VstarConn.Execute " Insert Into NewCaptions Values('33006','Leave Name','33006') "
VstarConn.Execute " Insert Into NewCaptions Values('33007','Please Wait...Updating Yearly Leaves','33007') "
VstarConn.Execute " Insert Into NewCaptions Values('33008','&Update','33008') "
VstarConn.Execute " Insert Into NewCaptions Values('33009','Yearly Leaves Already Updated ,Still Continue ?','33009') "
VstarConn.Execute " Insert Into NewCaptions Values('35001','Insert Memo Text','35001') "
VstarConn.Execute " Insert Into NewCaptions Values('35002','Memo','35002') "
VstarConn.Execute " Insert Into NewCaptions Values('35003','Ignore for','35003') "
VstarConn.Execute " Insert Into NewCaptions Values('36001','Monthly Process','36001') "
VstarConn.Execute " Insert Into NewCaptions Values('36002','Process for','36002') "
VstarConn.Execute " Insert Into NewCaptions Values('36003','All Employees','36003') "
VstarConn.Execute " Insert Into NewCaptions Values('36004','Selected Employees','36004') "
VstarConn.Execute " Insert Into NewCaptions Values('36005','Employees','36005') "
VstarConn.Execute " Insert Into NewCaptions Values('36006','Execute Late / Early rules','36006') "
VstarConn.Execute " Insert Into NewCaptions Values('36007','Consider Data from last month''s file','36007') "
VstarConn.Execute " Insert Into NewCaptions Values('36008','from the day','36008') "
VstarConn.Execute " Insert Into NewCaptions Values('36009','Onwards','36009') "
VstarConn.Execute " Insert Into NewCaptions Values('36010','&Process','36010') "
VstarConn.Execute " Insert Into NewCaptions Values('36011','Click this only for final Process.','36011') "
VstarConn.Execute " Insert Into NewCaptions Values('36012','If Processed again you may get false results.','36012') "
VstarConn.Execute " Insert Into NewCaptions Values('36013','Monthly Process Complete','36013') "
VstarConn.Execute " Insert Into NewCaptions Values('36014','     Please mention the day from which','36014') "
VstarConn.Execute " Insert Into NewCaptions Values('36015','the data from last month''s file has to be taken.','36015') "
VstarConn.Execute " Insert Into NewCaptions Values('36016',' Day cannot be greater than 31','36016') "
VstarConn.Execute " Insert Into NewCaptions Values('36017','Leave balance file for the year','36017') "
VstarConn.Execute " Insert Into NewCaptions Values('36018','Leave Information file for the year','36018') "
VstarConn.Execute " Insert Into NewCaptions Values('36019','Transaction file for the month','36019') "
VstarConn.Execute " Insert Into NewCaptions Values('36020','Difference between From Date and To Date cannot be more than 31','36020') "
VstarConn.Execute " Insert Into NewCaptions Values('36021','Difference between From Month and To Month cannot be more than 1','36021') "
VstarConn.Execute " Insert Into NewCaptions Values('36022',' does not exist in Leave Transaction File','36022') "
VstarConn.Execute " Insert Into NewCaptions Values('36023','Can not continue process.','36023') "
VstarConn.Execute " Insert Into NewCaptions Values('37001','Open Balance Entry','37001') "
VstarConn.Execute " Insert Into NewCaptions Values('37002','Opening on.','37002') "
VstarConn.Execute " Insert Into NewCaptions Values('37003','Please Select the Leave to be Added as Opening Balance','37003') "
VstarConn.Execute " Insert Into NewCaptions Values('37004','Opening Balance Leave(s) cannot be Added for 0 Number of days','37004') "
VstarConn.Execute " Insert Into NewCaptions Values('37005','Opening Balance Leave Days Must be Divisible by 0.50','37005') "
VstarConn.Execute " Insert Into NewCaptions Values('37006','Leaves already Added as Opening Balance on the above Selected Date','37006') "
VstarConn.Execute " Insert Into NewCaptions Values('40001','Reports','40001') "
VstarConn.Execute " Insert Into NewCaptions Values('40002','&Daily','40002') "
VstarConn.Execute " Insert Into NewCaptions Values('40003','Daily &Reports','40003') "
VstarConn.Execute " Insert Into NewCaptions Values('40004','Report for the date of','40004') "
VstarConn.Execute " Insert Into NewCaptions Values('40005','Shift Code','40005') "
VstarConn.Execute " Insert Into NewCaptions Values('40006','Physical Arrival','40006') "
VstarConn.Execute " Insert Into NewCaptions Values('40007','Absent','40007') "
VstarConn.Execute " Insert Into NewCaptions Values('40008','Continuous Absent','40008') "
VstarConn.Execute " Insert Into NewCaptions Values('40009','Late Arrival','40009') "
VstarConn.Execute " Insert Into NewCaptions Values('40010','Early Departure','40010') "
VstarConn.Execute " Insert Into NewCaptions Values('40011','Performance','40011') "
VstarConn.Execute " Insert Into NewCaptions Values('40012','Irregular','40012') "
VstarConn.Execute " Insert Into NewCaptions Values('40013','Entries','40013') "
VstarConn.Execute " Insert Into NewCaptions Values('40014','Shift Arrangement','40014') "
VstarConn.Execute " Insert Into NewCaptions Values('40015','Manpower','40015') "
VstarConn.Execute " Insert Into NewCaptions Values('40016','Out door duty','40016') "
VstarConn.Execute " Insert Into NewCaptions Values('40017','Summary','40017') "
VstarConn.Execute " Insert Into NewCaptions Values('40018','&Weekly','40018') "
VstarConn.Execute " Insert Into NewCaptions Values('40019','Weekly &Reports','40019') "
VstarConn.Execute " Insert Into NewCaptions Values('40020','Report for Week Starting From','40020') "
VstarConn.Execute " Insert Into NewCaptions Values('40021','Attendance','40021') "
VstarConn.Execute " Insert Into NewCaptions Values('40022','Shift Schedule','40022') "
VstarConn.Execute " Insert Into NewCaptions Values('40023','&Monthly','40023') "
VstarConn.Execute " Insert Into NewCaptions Values('40024','Monthly &Reports','40024') "
VstarConn.Execute " Insert Into NewCaptions Values('40025','Report for the month of','40025') "
VstarConn.Execute " Insert Into NewCaptions Values('40026','Muster Report','40026') "
VstarConn.Execute " Insert Into NewCaptions Values('40027','Monthly Present','40027') "
VstarConn.Execute " Insert Into NewCaptions Values('40028','Monthly Absent','40028') "
VstarConn.Execute " Insert Into NewCaptions Values('40029','Overtime Paid','40029') "
VstarConn.Execute " Insert Into NewCaptions Values('40030','Absent Memo','40030') "
VstarConn.Execute " Insert Into NewCaptions Values('40031','Late/Early/Absent','40031') "
VstarConn.Execute " Insert Into NewCaptions Values('40032','Leave Balance','40032') "
VstarConn.Execute " Insert Into NewCaptions Values('40033','Late Arrival Memo','40033') "
VstarConn.Execute " Insert Into NewCaptions Values('40034','Early Departure Memo','40034') "
VstarConn.Execute " Insert Into NewCaptions Values('40035','Leave Consumption','40035') "
VstarConn.Execute " Insert Into NewCaptions Values('40036','Total Lates','40036') "
VstarConn.Execute " Insert Into NewCaptions Values('40037','Total Earlys','40037') "
VstarConn.Execute " Insert Into NewCaptions Values('40038','WO on Holiday','40038') "
VstarConn.Execute " Insert Into NewCaptions Values('40039','&Yearly','40039') "
VstarConn.Execute " Insert Into NewCaptions Values('40040','Yearly &Reports','40040') "
VstarConn.Execute " Insert Into NewCaptions Values('40041','Report for the year of','40041') "
VstarConn.Execute " Insert Into NewCaptions Values('40042','Man Days','40042') "
VstarConn.Execute " Insert Into NewCaptions Values('40043','Present','40043') "
VstarConn.Execute " Insert Into NewCaptions Values('40044','Leave Information','40044') "
VstarConn.Execute " Insert Into NewCaptions Values('40045','M&asters','40045') "
VstarConn.Execute " Insert Into NewCaptions Values('40046','Master &Reports','40046') "
VstarConn.Execute " Insert Into NewCaptions Values('40047','Employee List','40047') "
VstarConn.Execute " Insert Into NewCaptions Values('40048','Employee Details','40048') "
VstarConn.Execute " Insert Into NewCaptions Values('40049','Left Employees','40049') "
VstarConn.Execute " Insert Into NewCaptions Values('40050','Rotational Shift','40050') "
VstarConn.Execute " Insert Into NewCaptions Values('40051','Holiday','40051') "
VstarConn.Execute " Insert Into NewCaptions Values('40052','P&eriodic','40052') "
VstarConn.Execute " Insert Into NewCaptions Values('40053','Periodic &Reports','40053') "
VstarConn.Execute " Insert Into NewCaptions Values('40054','Report for the period from','40054') "
VstarConn.Execute " Insert Into NewCaptions Values('40055','Reports Available for 30/31 Days only','40055') "
VstarConn.Execute " Insert Into NewCaptions Values('40056','Selectio&n','40056') "
VstarConn.Execute " Insert Into NewCaptions Values('40057','Employee','40057') "
VstarConn.Execute " Insert Into NewCaptions Values('40058','Group by','40058') "
VstarConn.Execute " Insert Into NewCaptions Values('40059','Department/Category','40059') "
VstarConn.Execute " Insert Into NewCaptions Values('40060','Start new page when group changes','40060') "
VstarConn.Execute " Insert Into NewCaptions Values('40061','Print Date && Time','40061') "
VstarConn.Execute " Insert Into NewCaptions Values('40062','Use 132 Column Dot matrix Printer','40062') "
VstarConn.Execute " Insert Into NewCaptions Values('40063','Prompt before Printing','40063') "
VstarConn.Execute " Insert Into NewCaptions Values('40064','Please confirm that your default Printer is set to 132 column Dot Matrix Printer.','40064') "
VstarConn.Execute " Insert Into NewCaptions Values('40065','Monthly Transactin File not found for the Month of','40065') "
VstarConn.Execute " Insert Into NewCaptions Values('40066','Cannot Send Reports through Email :: Refer Install->Parameters','40066') "
VstarConn.Execute " Insert Into NewCaptions Values('40067','Perform printing?','40067') "
VstarConn.Execute " Insert Into NewCaptions Values('40068','Please Enter the Date First','40068') "
VstarConn.Execute " Insert Into NewCaptions Values('40069','Daily Process Required. Continue ?','40069') "
VstarConn.Execute " Insert Into NewCaptions Values('40070','If Monthly process for selected month is not done,','40070') "
VstarConn.Execute " Insert Into NewCaptions Values('40071','please do it first.','40071') "
VstarConn.Execute " Insert Into NewCaptions Values('40072','Network Problem : Please Retry','40072') "
VstarConn.Execute " Insert Into NewCaptions Values('40073','The Papersize of the selected Printer is smaller than','40073') "
VstarConn.Execute " Insert Into NewCaptions Values('40074','      the required Papersize for the Report.','40074') "
VstarConn.Execute " Insert Into NewCaptions Values('40075','Operation is cancelled .','40075') "
VstarConn.Execute " Insert Into NewCaptions Values('40076','Shift File not found for the Month of','40076') "
VstarConn.Execute " Insert Into NewCaptions Values('40077','Month can not be empty','40077') "
VstarConn.Execute " Insert Into NewCaptions Values('40078','Year can not be empty','40078') "
VstarConn.Execute " Insert Into NewCaptions Values('40079','Period should not be more than 31','40079') "
VstarConn.Execute " Insert Into NewCaptions Values('40080','Period should not be for more than 2 Months','40080') "
VstarConn.Execute " Insert Into NewCaptions Values('40081','Yearly Leave Transaction File Not Found','40081') "
VstarConn.Execute " Insert Into NewCaptions Values('40082','Yearly Leave Information File Not Found','40082') "
VstarConn.Execute " Insert Into NewCaptions Values('40083','Mail Address not found for','40083') "
VstarConn.Execute " Insert Into NewCaptions Values('40084','Meal Allowance','40084') "
VstarConn.Execute " Insert Into NewCaptions Values('40085','Punches Report','40085') "
VstarConn.Execute " Insert Into NewCaptions Values('40086','Permission cards','40086') "
VstarConn.Execute " Insert Into NewCaptions Values('41001','Rotation Master','41001') "
VstarConn.Execute " Insert Into NewCaptions Values('41002','Rotation Code','41002') "
VstarConn.Execute " Insert Into NewCaptions Values('41003','Shift Rotates','41003') "
VstarConn.Execute " Insert Into NewCaptions Values('41004','Only after specified number of days','41004') "
VstarConn.Execute " Insert Into NewCaptions Values('41005','On the following dates of every month','41005') "
VstarConn.Execute " Insert Into NewCaptions Values('41006','On following week days(SUN..SAT)','41006') "
VstarConn.Execute " Insert Into NewCaptions Values('41007','The shift changes from one to another','41007') "
VstarConn.Execute " Insert Into NewCaptions Values('41008','Rotation','41008') "
VstarConn.Execute " Insert Into NewCaptions Values('41009','Shift Rotation','41009') "
VstarConn.Execute " Insert Into NewCaptions Values('41010','Rotation Days/Dates','41010') "
VstarConn.Execute " Insert Into NewCaptions Values('41011','Rotation Code Already Exists','41011') "
VstarConn.Execute " Insert Into NewCaptions Values('41012','Rotation Name cannot be Blank','41012') "
VstarConn.Execute " Insert Into NewCaptions Values('41013','Days are not Specified , Please Do It','41013') "
VstarConn.Execute " Insert Into NewCaptions Values('41014','Dates are not Specified , Please Do It','41014') "
VstarConn.Execute " Insert Into NewCaptions Values('41015','Week Days are not Specified , Please Do It','41015') "
VstarConn.Execute " Insert Into NewCaptions Values('41016','Shifts are not Specified , Please Do It','41016') "
VstarConn.Execute " Insert Into NewCaptions Values('41017','Rotation with Same Code Already Exists','41017') "
VstarConn.Execute " Insert Into NewCaptions Values('41018','Click to Select the Number of Days','41018') "
VstarConn.Execute " Insert Into NewCaptions Values('41019','Click to Select the Dates of Every Month','41019') "

VstarConn.Execute " Insert Into NewCaptions Values('41020','Click to Select the Week','41020') "
VstarConn.Execute " Insert Into NewCaptions Values('41021','Click to Select the Shiffts','41021') "
VstarConn.Execute " Insert Into NewCaptions Values('41022','Rotation Code cannot be blank','41022') "
VstarConn.Execute " Insert Into NewCaptions Values('42001','Only after specified number of days','42001') "
VstarConn.Execute " Insert Into NewCaptions Values('43001','Select Week Days','43001') "
VstarConn.Execute " Insert Into NewCaptions Values('44001','Late /Early Rules','44001') "
VstarConn.Execute " Insert Into NewCaptions Values('44002','Select C&ategory','44002') "
VstarConn.Execute " Insert Into NewCaptions Values('44003','Late Deductions','44003') "
VstarConn.Execute " Insert Into NewCaptions Values('44004','Total Late Allowed in a Month','44004') "
VstarConn.Execute " Insert Into NewCaptions Values('44005','Cut','44005') "
VstarConn.Execute " Insert Into NewCaptions Values('44006','Day for every','44006') "
VstarConn.Execute " Insert Into NewCaptions Values('44007','&Deduct From','44007') "
VstarConn.Execute " Insert Into NewCaptions Values('44008','Paid Days','44008') "
VstarConn.Execute " Insert Into NewCaptions Values('44009','Leaves','44009') "
VstarConn.Execute " Insert Into NewCaptions Values('44010','1st Preference','44010') "
VstarConn.Execute " Insert Into NewCaptions Values('44011','2nd Preference','44011') "
VstarConn.Execute " Insert Into NewCaptions Values('44012','3rd Preference','44012') "
VstarConn.Execute " Insert Into NewCaptions Values('44013','Early Deductions','44013') "
VstarConn.Execute " Insert Into NewCaptions Values('44014','Total Early Allowed in Month','44014') "
VstarConn.Execute " Insert Into NewCaptions Values('44015','&Reset','44015') "
VstarConn.Execute " Insert Into NewCaptions Values('44016','Are You Sure to Reset the Rules','44016') "
VstarConn.Execute " Insert Into NewCaptions Values('44017','Days To be Cut Must be Divisible by 0.50','44017') "
VstarConn.Execute " Insert Into NewCaptions Values('44018','Fractional Number not Allowed','44018') "
VstarConn.Execute " Insert Into NewCaptions Values('44019','Please Select atleast 1 Leave','44019') "
VstarConn.Execute " Insert Into NewCaptions Values('44020','Please Select Leave for the Second Preference','44020') "
VstarConn.Execute " Insert Into NewCaptions Values('44021','Please Select Leave for the First Preference','44021') "
VstarConn.Execute " Insert Into NewCaptions Values('45001','Shift schedule for','45001') "
VstarConn.Execute " Insert Into NewCaptions Values('45002','1st Week','45002') "
VstarConn.Execute " Insert Into NewCaptions Values('45003','2nd Week','45003') "
VstarConn.Execute " Insert Into NewCaptions Values('45004','3rd Week','45004') "
VstarConn.Execute " Insert Into NewCaptions Values('45005','4th Week','45005') "
VstarConn.Execute " Insert Into NewCaptions Values('45006','5th Week','45006') "
VstarConn.Execute " Insert Into NewCaptions Values('45007','&Period','45007') "
VstarConn.Execute " Insert Into NewCaptions Values('45008','M&aster','45008') "
VstarConn.Execute " Insert Into NewCaptions Values('45009','MON','45009') "
VstarConn.Execute " Insert Into NewCaptions Values('45010','TUE','45010') "
VstarConn.Execute " Insert Into NewCaptions Values('45011','WED','45011') "
VstarConn.Execute " Insert Into NewCaptions Values('45012','THU','45012') "
VstarConn.Execute " Insert Into NewCaptions Values('45013','FRI','45013') "
VstarConn.Execute " Insert Into NewCaptions Values('45014','SAT','45014') "
VstarConn.Execute " Insert Into NewCaptions Values('45015','SUN','45015') "
VstarConn.Execute " Insert Into NewCaptions Values('45016','Monthly Shift File not Found for the Month of','45016') "
VstarConn.Execute " Insert Into NewCaptions Values('45017','Shift for the Employee','45017') "
VstarConn.Execute " Insert Into NewCaptions Values('45018',' not Yet Created :: Please Create it','45018') "
VstarConn.Execute " Insert Into NewCaptions Values('46001','Select Dat File','46001') "
VstarConn.Execute " Insert Into NewCaptions Values('46002','&Files','46002') "
VstarConn.Execute " Insert Into NewCaptions Values('46003','Fold&ers','46003') "
VstarConn.Execute " Insert Into NewCaptions Values('46004','&Drives','46004') "
VstarConn.Execute " Insert Into NewCaptions Values('47001','Select Shifts','47001') "
VstarConn.Execute " Insert Into NewCaptions Values('47002','Start','47002') "
VstarConn.Execute " Insert Into NewCaptions Values('47003','End','47003') "
VstarConn.Execute " Insert Into NewCaptions Values('47004','Night','47004') "
VstarConn.Execute " Insert Into NewCaptions Values('47005','&Reset','47005') "
VstarConn.Execute " Insert Into NewCaptions Values('48001','Shift Master','48001') "
VstarConn.Execute " Insert Into NewCaptions Values('48002','Shift Code','48002') "
VstarConn.Execute " Insert Into NewCaptions Values('48003','This is a night shift','48003') "
VstarConn.Execute " Insert Into NewCaptions Values('48004','Deduct Break Hrs from Shift hrs','48004') "
VstarConn.Execute " Insert Into NewCaptions Values('48005','Shift Time','48005') "
VstarConn.Execute " Insert Into NewCaptions Values('48006','Shift starts at','48006') "
VstarConn.Execute " Insert Into NewCaptions Values('48007','First half ends at','48007') "
VstarConn.Execute " Insert Into NewCaptions Values('48008','Second half starts at','48008') "
VstarConn.Execute " Insert Into NewCaptions Values('48009','Shift Ends at','48009') "
VstarConn.Execute " Insert Into NewCaptions Values('48010','Total shift time','48010') "
VstarConn.Execute " Insert Into NewCaptions Values('48011','Break Periods','48011') "
VstarConn.Execute " Insert Into NewCaptions Values('48012','Starts at','48012') "
VstarConn.Execute " Insert Into NewCaptions Values('48013','Ends at','48013') "
VstarConn.Execute " Insert Into NewCaptions Values('48014','Break','48014') "
VstarConn.Execute " Insert Into NewCaptions Values('48015','First Break','48015') "
VstarConn.Execute " Insert Into NewCaptions Values('48016','Second Break','48016') "
VstarConn.Execute " Insert Into NewCaptions Values('48017','Third Break','48017') "
VstarConn.Execute " Insert Into NewCaptions Values('48018','Shift Code cannot be blank','48018') "
VstarConn.Execute " Insert Into NewCaptions Values('48019','Shift Already Exists','48019') "
VstarConn.Execute " Insert Into NewCaptions Values('48020','Shift Name cannot be blank','48020') "
VstarConn.Execute " Insert Into NewCaptions Values('48021','Second Shift Start time cannot Be Less than First Shift End Time','48021') "
VstarConn.Execute " Insert Into NewCaptions Values('48022','First Break End Time cannot be Less than First Break Start Time','48022') "
VstarConn.Execute " Insert Into NewCaptions Values('48023','Second Break End Time cannot be Less than Second Break Start Time','48023') "
VstarConn.Execute " Insert Into NewCaptions Values('48024','Third Break End Time cannot be Less than Third Break Start Time','48024') "
VstarConn.Execute " Insert Into NewCaptions Values('48025','Second Break Start Time Cannot be Less than First Break End Time','48025') "
VstarConn.Execute " Insert Into NewCaptions Values('48026','Third Break Start Time Cannot be Less than Second Break End Time','48026') "
VstarConn.Execute " Insert Into NewCaptions Values('48027','Time Should be Greater than Shift Start Time','48027') "
VstarConn.Execute " Insert Into NewCaptions Values('48028','Time Should be Less than Shift End Time','48028') "
VstarConn.Execute " Insert Into NewCaptions Values('48029','Shift with Same Code Already Exists','48029') "
VstarConn.Execute " Insert Into NewCaptions Values('48030','Shift not Found','48030') "
VstarConn.Execute " Insert Into NewCaptions Values('48031','Are you Sure to Delete this Record','48031') "
VstarConn.Execute " Insert Into NewCaptions Values('49001','Monthly shift creation','49001') "
VstarConn.Execute " Insert Into NewCaptions Values('49004','Shift Date','49004') "
VstarConn.Execute " Insert Into NewCaptions Values('49005',':: Please Wait...Processing Shifts','49005') "
VstarConn.Execute " Insert Into NewCaptions Values('49006','Shift file processed successfully for the month of','49006') "
VstarConn.Execute " Insert Into NewCaptions Values('49007','Shift file Already Exists for the Month of','49007') "
VstarConn.Execute " Insert Into NewCaptions Values('49008',' Do You Wish to Overwrite it','49008') "
VstarConn.Execute " Insert Into NewCaptions Values('49009','Employee Selection','49009') "
VstarConn.Execute " Insert Into NewCaptions Values('49010','Click to Toggle Selection','49010') "
VstarConn.Execute " Insert Into NewCaptions Values('49011','Create monthly shift schedule','49011') "
VstarConn.Execute " Insert Into NewCaptions Values('50001','Select Shift','50001') "
VstarConn.Execute " Insert Into NewCaptions Values('50002','Start','50002') "
VstarConn.Execute " Insert Into NewCaptions Values('50003','End','50003') "
VstarConn.Execute " Insert Into NewCaptions Values('50004','Night','50004') "
VstarConn.Execute " Insert Into NewCaptions Values('50005','Please Select the Shift','50005') "
VstarConn.Execute " Insert Into NewCaptions Values('52001','Login','52001') "
VstarConn.Execute " Insert Into NewCaptions Values('52002','Password','52002') "
VstarConn.Execute " Insert Into NewCaptions Values('52003','User Name','52003') "
VstarConn.Execute " Insert Into NewCaptions Values('52004','Exclusive Access','52004') "
VstarConn.Execute " Insert Into NewCaptions Values('52005','User Name cannot be Blank','52005') "
VstarConn.Execute " Insert Into NewCaptions Values('52006','There is an Exclusive Lock by the Administrator :: Cannot Start the Application','52006') "
VstarConn.Execute " Insert Into NewCaptions Values('52007','Invalid Username or Password :: Try Again','52007') "
VstarConn.Execute " Insert Into NewCaptions Values('53001','&Install','53001') "
VstarConn.Execute " Insert Into NewCaptions Values('53002','Pa&rameter','53002') "
VstarConn.Execute " Insert Into NewCaptions Values('53003','&Shift','53003') "
VstarConn.Execute " Insert Into NewCaptions Values('53004','C&ompany','53004') "
VstarConn.Execute " Insert Into NewCaptions Values('53005','&Category','53005') "
VstarConn.Execute " Insert Into NewCaptions Values('53006','&Leave','53006') "
VstarConn.Execute " Insert Into NewCaptions Values('53007','&Yearly Leaves','53007') "
VstarConn.Execute " Insert Into NewCaptions Values('53008','&Update','53008') "
VstarConn.Execute " Insert Into NewCaptions Values('53009','&Rules','53009') "
VstarConn.Execute " Insert Into NewCaptions Values('53010','Login as &Different User','53010') "
VstarConn.Execute " Insert Into NewCaptions Values('53011','&Updation','53011') "
VstarConn.Execute " Insert Into NewCaptions Values('53012','&Department Master','53012') "
VstarConn.Execute " Insert Into NewCaptions Values('53013','&Group Master','53013') "
VstarConn.Execute " Insert Into NewCaptions Values('53014','&Rotation Shift Master','53014') "
VstarConn.Execute " Insert Into NewCaptions Values('53015','&Holidays','53015') "
VstarConn.Execute " Insert Into NewCaptions Values('53016','Ho&liday Master','53016') "
VstarConn.Execute " Insert Into NewCaptions Values('53017','&Declare Holiday','53017') "
VstarConn.Execute " Insert Into NewCaptions Values('53018','&Employee Master','53018') "
VstarConn.Execute " Insert Into NewCaptions Values('53019','&Shift Schedule','53019') "
VstarConn.Execute " Insert Into NewCaptions Values('53020','&Schedule Master','53020') "
VstarConn.Execute " Insert Into NewCaptions Values('53021','&Create Schedule','53021') "
VstarConn.Execute " Insert Into NewCaptions Values('53022','Chan&ge Schedule','53022') "
VstarConn.Execute " Insert Into NewCaptions Values('53023','&Leaves Transaction','53023') "
VstarConn.Execute " Insert Into NewCaptions Values('53024','&Opening','53024') "
VstarConn.Execute " Insert Into NewCaptions Values('53025','&Credit','53025') "
VstarConn.Execute " Insert Into NewCaptions Values('53026','&Encash','53026') "
VstarConn.Execute " Insert Into NewCaptions Values('53027','&Avail','53027') "
VstarConn.Execute " Insert Into NewCaptions Values('53028','Lo&st Entry','53028') "
VstarConn.Execute " Insert Into NewCaptions Values('53029','C&orrection','53029') "
VstarConn.Execute " Insert Into NewCaptions Values('53030','Edit &Paid Days','53030') "
VstarConn.Execute " Insert Into NewCaptions Values('53031','&Process','53031') "
VstarConn.Execute " Insert Into NewCaptions Values('53032','&Daily','53032') "
VstarConn.Execute " Insert Into NewCaptions Values('53033','&Monthly','53033') "
VstarConn.Execute " Insert Into NewCaptions Values('53034','&Report','53034') "
VstarConn.Execute " Insert Into NewCaptions Values('53035','&Reports','53035') "
VstarConn.Execute " Insert Into NewCaptions Values('53036','U&tility','53036') "
VstarConn.Execute " Insert Into NewCaptions Values('53037','Change &Password','53037') "
VstarConn.Execute " Insert Into NewCaptions Values('53038','&Backup and Restore','53038') "
VstarConn.Execute " Insert Into NewCaptions Values('53039','&Reset Locks','53039') "
VstarConn.Execute " Insert Into NewCaptions Values('53040','Compact &Database','53040') "
VstarConn.Execute " Insert Into NewCaptions Values('53041','&Version','53041') "
VstarConn.Execute " Insert Into NewCaptions Values('53042','&About','53042') "
VstarConn.Execute " Insert Into NewCaptions Values('53043','Support','53043') "
VstarConn.Execute " Insert Into NewCaptions Values('53044','User &Accounts','53044') "
VstarConn.Execute " Insert Into NewCaptions Values('53045','Monthly Data is  in Process : Please Try after Some Time','53045') "
VstarConn.Execute " Insert Into NewCaptions Values('53046','Yearly Leave Files are in Process :: Please Try after some Time','53046') "
VstarConn.Execute " Insert Into NewCaptions Values('53047','Data is  in Process : Please Try after Some Time','53047') "
VstarConn.Execute " Insert Into NewCaptions Values('53048','&Create','53048') "
VstarConn.Execute " Insert Into NewCaptions Values('53049','&Language','53049') "
VstarConn.Execute " Insert Into NewCaptions Values('53050','&Switch Language','53050') "
VstarConn.Execute " Insert Into NewCaptions Values('53051','&Translate','53051') "
VstarConn.Execute " Insert Into NewCaptions Values('53052','&Report Captions','53052') "
VstarConn.Execute " Insert Into NewCaptions Values('53053','&General Captions','53053') "
VstarConn.Execute " Insert Into NewCaptions Values('53054','&Change Menu Captions','53054') "
VstarConn.Execute " Insert Into NewCaptions Values('53055','E&xit','53055') "
VstarConn.Execute " Insert Into NewCaptions Values('53056','&Daily Process','53056') "
VstarConn.Execute " Insert Into NewCaptions Values('53057','&Monthly Process','53057') "
VstarConn.Execute " Insert Into NewCaptions Values('53058','O&T Rules','53058') "
VstarConn.Execute " Insert Into NewCaptions Values('53059','C&O Rules','53059') "
VstarConn.Execute " Insert Into NewCaptions Values('53060','C&hange User Password','53060') "
VstarConn.Execute " Insert Into NewCaptions Values('53061','OT &Authorization','53061') "
VstarConn.Execute " Insert Into NewCaptions Values('53062','Loc&ation Master','53062') "
VstarConn.Execute " Insert Into NewCaptions Values('53063','&Set Employee Details','53063') "
VstarConn.Execute " Insert Into NewCaptions Values('53064','&Division Master','53064') "
VstarConn.Execute " Insert Into NewCaptions Values('53065','&Export Data','53065') "
'' apoorva
VstarConn.Execute " Insert Into NewCaptions Values('53066','&Dos Reports','53066') "
VstarConn.Execute " Insert Into NewCaptions Values('54001','Installation Parameters','54001') "
VstarConn.Execute " Insert Into NewCaptions Values('54002','Parameters','54002') "
VstarConn.Execute " Insert Into NewCaptions Values('54003','General','54003') "
VstarConn.Execute " Insert Into NewCaptions Values('54004','Current year is','54004') "
VstarConn.Execute " Insert Into NewCaptions Values('54005','Employee code size is','54005') "
VstarConn.Execute " Insert Into NewCaptions Values('54006','Punching card size is','54006') "
VstarConn.Execute " Insert Into NewCaptions Values('54007','Year Starting from','54007') "
VstarConn.Execute " Insert Into NewCaptions Values('54008','Week begins on','54008') "
VstarConn.Execute " Insert Into NewCaptions Values('54009','Employee can work after the shift for next','54009') "
VstarConn.Execute " Insert Into NewCaptions Values('54010','Select Path for .DAT File','54010') "
VstarConn.Execute " Insert Into NewCaptions Values('54011','Ignore next punch from  the previous punch till','54011') "
VstarConn.Execute " Insert Into NewCaptions Values('54012','Allow Employee to be posted to next shift if late by','54012') "
VstarConn.Execute " Insert Into NewCaptions Values('54013','Allow Employee to be posted to Previous shift if Early by','54013') "
VstarConn.Execute " Insert Into NewCaptions Values('54014','Details','54014') "
VstarConn.Execute " Insert Into NewCaptions Values('54015','Permission Cards','54015') "
VstarConn.Execute " Insert Into NewCaptions Values('54016','Use Permission cards','54016') "
VstarConn.Execute " Insert Into NewCaptions Values('54017','Starting number','54017') "
VstarConn.Execute " Insert Into NewCaptions Values('54018','Ending number','54018') "
VstarConn.Execute " Insert Into NewCaptions Values('54019','Card number','54019') "
VstarConn.Execute " Insert Into NewCaptions Values('54020','Late Coming','54020') "
VstarConn.Execute " Insert Into NewCaptions Values('54021','Early Going','54021') "
VstarConn.Execute " Insert Into NewCaptions Values('54022','Late bus','54022') "
VstarConn.Execute " Insert Into NewCaptions Values('54023','Official Duty','54023') "
VstarConn.Execute " Insert Into NewCaptions Values('54024','OverTime','54024') "
VstarConn.Execute " Insert Into NewCaptions Values('54025','Rates','54025') "
VstarConn.Execute " Insert Into NewCaptions Values('54026','On a Holiday, Calculate @','54026') "
VstarConn.Execute " Insert Into NewCaptions Values('54027','On a Week Off,Calculate @','54027') "
VstarConn.Execute " Insert Into NewCaptions Values('54028','On a Working Day,Calculate @','54028') "
VstarConn.Execute " Insert Into NewCaptions Values('54029','times, the  total work hours for that day','54029') "
VstarConn.Execute " Insert Into NewCaptions Values('54030','Deduct Late / Early Hours from OverTime','54030') "

VstarConn.Execute " Insert Into NewCaptions Values('54031','Roundoff Over Time Hours','54031') "
VstarConn.Execute " Insert Into NewCaptions Values('54032','&Back','54032') "
VstarConn.Execute " Insert Into NewCaptions Values('54033','&Next','54033') "
VstarConn.Execute " Insert Into NewCaptions Values('54034','Select Path for .DAT File','54034') "
VstarConn.Execute " Insert Into NewCaptions Values('54035','digit','54035') "
VstarConn.Execute " Insert Into NewCaptions Values('54036','Send Reports using Email','54036') "
VstarConn.Execute " Insert Into NewCaptions Values('54037','Overtime Rounding','54037') "
VstarConn.Execute " Insert Into NewCaptions Values('54038','Decimal Value (in Minutes)','54038') "
VstarConn.Execute " Insert Into NewCaptions Values('54039','OT Round off Pattern','54039') "
VstarConn.Execute " Insert Into NewCaptions Values('54040','Round off to','54040') "
VstarConn.Execute " Insert Into NewCaptions Values('54041','More than','54041') "
VstarConn.Execute " Insert Into NewCaptions Values('54042','Apply Salary Cut-Off Date','54042') "
VstarConn.Execute " Insert Into NewCaptions Values('54043','Cut-Off Date','54043') "
VstarConn.Execute " Insert Into NewCaptions Values('54044','Application Date Format :','54044') "
VstarConn.Execute " Insert Into NewCaptions Values('54045','Size Once Increased cannot be Decreased','54045') "
VstarConn.Execute " Insert Into NewCaptions Values('54046','Warning : Size should not be decreased','54046') "
VstarConn.Execute " Insert Into NewCaptions Values('54047','This value should be greater than','54047') "
VstarConn.Execute " Insert Into NewCaptions Values('54048','Cut-Off Date Cannot be Greater than 31','54048') "
VstarConn.Execute " Insert Into NewCaptions Values('54049','Please enter appropriate Values in the OD Card','54049') "
VstarConn.Execute " Insert Into NewCaptions Values('54050','Please enter appropriate Values in the Late Card','54050') "
VstarConn.Execute " Insert Into NewCaptions Values('54051','Please enter appropriate Values in the Late Bus Card','54051') "
VstarConn.Execute " Insert Into NewCaptions Values('54052','Please enter appropriate Values in the Early  Card','54052') "
VstarConn.Execute " Insert Into NewCaptions Values('54053','Please enter appropriate Values in the Start Card','54053') "
VstarConn.Execute " Insert Into NewCaptions Values('54054','Please enter appropriate Values in the End Card','54054') "
VstarConn.Execute " Insert Into NewCaptions Values('54055','OFF DUTY card no should be in the range of','54055') "
VstarConn.Execute " Insert Into NewCaptions Values('54056','LATE card no should be in the range of','54056') "
VstarConn.Execute " Insert Into NewCaptions Values('54057','EARLY card no should be in the range of','54057') "
VstarConn.Execute " Insert Into NewCaptions Values('54058','LATE BUS card no should be in the range of','54058') "
VstarConn.Execute " Insert Into NewCaptions Values('54059','Two cards are given same number. Cannot save the parameters.','54059') "
VstarConn.Execute " Insert Into NewCaptions Values('54060','If Salary Generation is after Month Completion then keep CutOff Day as 0','54060') "
VstarConn.Execute " Insert Into NewCaptions Values('54061','Maximum value can be 10.00','54061') "
VstarConn.Execute " Insert Into NewCaptions Values('56001','User Accounts','56001') "
VstarConn.Execute " Insert Into NewCaptions Values('56002','List','56002') "
VstarConn.Execute " Insert Into NewCaptions Values('56003','Details','56003') "
VstarConn.Execute " Insert Into NewCaptions Values('56004','Password / UserName','56004') "
VstarConn.Execute " Insert Into NewCaptions Values('56005','Master File Rights','56005') "
VstarConn.Execute " Insert Into NewCaptions Values('56006','Other Rights','56006') "
VstarConn.Execute " Insert Into NewCaptions Values('56007','Leave Transaction','56007') "
VstarConn.Execute " Insert Into NewCaptions Values('56008','Select/Unselect All','56008') "
VstarConn.Execute " Insert Into NewCaptions Values('56009','Parameter','56009') "
VstarConn.Execute " Insert Into NewCaptions Values('56010','Process','56010') "
VstarConn.Execute " Insert Into NewCaptions Values('56011','Yearly Leaves','56011') "
VstarConn.Execute " Insert Into NewCaptions Values('56012','Edit','56012') "
VstarConn.Execute " Insert Into NewCaptions Values('56013','Daily','56013') "
VstarConn.Execute " Insert Into NewCaptions Values('56014','Monthly','56014') "
VstarConn.Execute " Insert Into NewCaptions Values('56015','Create','56015') "
VstarConn.Execute " Insert Into NewCaptions Values('56016','Update','56016') "
VstarConn.Execute " Insert Into NewCaptions Values('56017','Change','56017') "
VstarConn.Execute " Insert Into NewCaptions Values('56018','Record','56018') "
VstarConn.Execute " Insert Into NewCaptions Values('56019','On duty','56019') "
VstarConn.Execute " Insert Into NewCaptions Values('56020','Off Duty','56020') "
VstarConn.Execute " Insert Into NewCaptions Values('56021','OT','56021') "
VstarConn.Execute " Insert Into NewCaptions Values('56022','Time','56022') "
VstarConn.Execute " Insert Into NewCaptions Values('56023','Add/Edit/Delete of Login Users','56023') "
VstarConn.Execute " Insert Into NewCaptions Values('56024','Backup','56024') "
VstarConn.Execute " Insert Into NewCaptions Values('56025','Restore','56025') "
VstarConn.Execute " Insert Into NewCaptions Values('56026','Login users','56026') "
VstarConn.Execute " Insert Into NewCaptions Values('56027','Correction','56027') "
VstarConn.Execute " Insert Into NewCaptions Values('56028','Shift Schedule','56028') "
VstarConn.Execute " Insert Into NewCaptions Values('56029','Paid Days','56029') "
VstarConn.Execute " Insert Into NewCaptions Values('56030','User Code','56030') "
VstarConn.Execute " Insert Into NewCaptions Values('56031','Password','56031') "
VstarConn.Execute " Insert Into NewCaptions Values('56032','Credit','56032') "
VstarConn.Execute " Insert Into NewCaptions Values('56033','Encash','56033') "
VstarConn.Execute " Insert Into NewCaptions Values('56034','Avail','56034') "
VstarConn.Execute " Insert Into NewCaptions Values('56035','Add','56035') "
VstarConn.Execute " Insert Into NewCaptions Values('56036','Delete','56036') "
VstarConn.Execute " Insert Into NewCaptions Values('56037','Opening','56037') "
VstarConn.Execute " Insert Into NewCaptions Values('56038','Rules','56038') "
VstarConn.Execute " Insert Into NewCaptions Values('56039','Compact Database','56039') "
VstarConn.Execute " Insert Into NewCaptions Values('56040','Permission','56040') "
VstarConn.Execute " Insert Into NewCaptions Values('56041','Maximum users allowed :','56041') "
VstarConn.Execute " Insert Into NewCaptions Values('56042','User Name cannot be blank.','56042') "
VstarConn.Execute " Insert Into NewCaptions Values('56043','Password cannot be blank.','56043') "
VstarConn.Execute " Insert Into NewCaptions Values('56044','Encryption Error :: Try Another Password','56044') "
VstarConn.Execute " Insert Into NewCaptions Values('56045','Unable to change the Password','56045') "
VstarConn.Execute " Insert Into NewCaptions Values('56046','Error changing the Password','56046') "
VstarConn.Execute " Insert Into NewCaptions Values('56047','The User Cannot Delete Himself','56047') "
VstarConn.Execute " Insert Into NewCaptions Values('56048','Are You Sure To Delete User','56048') "
VstarConn.Execute " Insert Into NewCaptions Values('56049','Please Select the User','56049') "
VstarConn.Execute " Insert Into NewCaptions Values('56050','The User Already Exists','56050') "
VstarConn.Execute " Insert Into NewCaptions Values('56051','No Records Obtained From User Master','56051') "
VstarConn.Execute " Insert Into NewCaptions Values('56052','Re-Start the Application','56052') "
VstarConn.Execute " Insert Into NewCaptions Values('56053','You are about to change your login password','56053') "
VstarConn.Execute " Insert Into NewCaptions Values('56054','Do you want to proceed ?','56054') "
VstarConn.Execute " Insert Into NewCaptions Values('56055','Your Login Password has been changed successfully','56055') "
VstarConn.Execute " Insert Into NewCaptions Values('56056','Menu Items','56056') "
VstarConn.Execute " Insert Into NewCaptions Values('56057','User Name','56057') "
VstarConn.Execute " Insert Into NewCaptions Values('56058','BackUp and Restore','56058') "
VstarConn.Execute " Insert Into NewCaptions Values('56059','Press Spacebar or Double Click to Toggle Rights','56059') "
VstarConn.Execute " Insert Into NewCaptions Values('57001','Visual star Version','57001') "
VstarConn.Execute " Insert Into NewCaptions Values('57002','Version information','57002') "
VstarConn.Execute " Insert Into NewCaptions Values('57003','Comments','57003') "
VstarConn.Execute " Insert Into NewCaptions Values('57004','Company Name','57004') "
VstarConn.Execute " Insert Into NewCaptions Values('57005','File Description','57005') "
VstarConn.Execute " Insert Into NewCaptions Values('57006','File Version','57006') "
VstarConn.Execute " Insert Into NewCaptions Values('57007','Internal Name','57007') "
VstarConn.Execute " Insert Into NewCaptions Values('57008','Legal copyright','57008') "
VstarConn.Execute " Insert Into NewCaptions Values('57009','Legal trademarks','57009') "
VstarConn.Execute " Insert Into NewCaptions Values('57010','Original Filename','57010') "
VstarConn.Execute " Insert Into NewCaptions Values('57011','Product Name','57011') "
VstarConn.Execute " Insert Into NewCaptions Values('57012','Product Version','57012') "
VstarConn.Execute " Insert Into NewCaptions Values('57013','Special Build for','57013') "
VstarConn.Execute " Insert Into NewCaptions Values('58001','Yearly File Creation','58001') "
VstarConn.Execute " Insert Into NewCaptions Values('58002','Yearly file creation for the selected Year','58002') "
VstarConn.Execute " Insert Into NewCaptions Values('58003','Use this option :','58003') "
VstarConn.Execute " Insert Into NewCaptions Values('58004','=> In the beginning of the year','58004') "
VstarConn.Execute " Insert Into NewCaptions Values('58005','=> When new type of leave is Added','58005') "
VstarConn.Execute " Insert Into NewCaptions Values('58006','=> When existing leave is Edited or Deleted','58006') "
VstarConn.Execute " Insert Into NewCaptions Values('58007','Instructions','58007') "
VstarConn.Execute " Insert Into NewCaptions Values('58008','&Create','58008') "
VstarConn.Execute " Insert Into NewCaptions Values('58009','File :','58009') "
VstarConn.Execute " Insert Into NewCaptions Values('58010','Already exists, You want to overwrite ?','58010') "
VstarConn.Execute " Insert Into NewCaptions Values('58011','Yearly Leave Files Creation finished','58011') "
VstarConn.Execute " Insert Into NewCaptions Values('59001','OT Rule No.','59001') "
VstarConn.Execute " Insert Into NewCaptions Values('59002','Give OT On','59002') "
VstarConn.Execute " Insert Into NewCaptions Values('59003','More than','59003') "
VstarConn.Execute " Insert Into NewCaptions Values('59004','Apply deduction on Weekoff','59004') "
VstarConn.Execute " Insert Into NewCaptions Values('59005','Apply deduction on Holiday','59005') "
VstarConn.Execute " Insert Into NewCaptions Values('59006','','59006') "
VstarConn.Execute " Insert Into NewCaptions Values('59007','Authorized by default','59007') "
VstarConn.Execute " Insert Into NewCaptions Values('59008','Maximum OT can be upto','59008') "
VstarConn.Execute " Insert Into NewCaptions Values('59009','Late - Early deductions','59009') "
VstarConn.Execute " Insert Into NewCaptions Values('59010','Deduct Late Hours from OT','59010') "
VstarConn.Execute " Insert Into NewCaptions Values('59011','Deduct Early Hours from OT','59011') "
VstarConn.Execute " Insert Into NewCaptions Values('59012','times, total work hours of that day','59012') "
VstarConn.Execute " Insert Into NewCaptions Values('59013','OT Rule cannot be empty.','59013') "
VstarConn.Execute " Insert Into NewCaptions Values('59014','OT Rule already exists.','59014') "
VstarConn.Execute " Insert Into NewCaptions Values('59015','OT Description cannot be empty.','59015') "
VstarConn.Execute " Insert Into NewCaptions Values('59016','To timings cannot be less than Deduct timings.','59016') "
VstarConn.Execute " Insert Into NewCaptions Values('59017','Please enter To timings.','59017') "
VstarConn.Execute " Insert Into NewCaptions Values('59018','OT Rule not found.','59018') "
VstarConn.Execute " Insert Into NewCaptions Values('59019','Please Select the OT Rule.','59019') "
VstarConn.Execute " Insert Into NewCaptions Values('59020','Weekdays @','59020') "
VstarConn.Execute " Insert Into NewCaptions Values('59021','Weekoffs @','59021') "
VstarConn.Execute " Insert Into NewCaptions Values('59022','Holidays @','59022') "
VstarConn.Execute " Insert Into NewCaptions Values('59023','Deduct the following hours from the Basic OT','59023') "
VstarConn.Execute " Insert Into NewCaptions Values('59024','Basic OT will be calculated after the LATE-EARLY calculations specified in category Master','59024') "
VstarConn.Execute " Insert Into NewCaptions Values('59025','Deduct specified','59025') "
VstarConn.Execute " Insert Into NewCaptions Values('59026','--OR--','59026') "
VstarConn.Execute " Insert Into NewCaptions Values('59027','Deduct all OT','59027') "
VstarConn.Execute " Insert Into NewCaptions Values('59028','Round-Off OT','59028') "
VstarConn.Execute " Insert Into NewCaptions Values('59029','While Rounding OT only MINUTES part will be rounded , leaving the hours part as it is','59059') "
VstarConn.Execute " Insert Into NewCaptions Values('59030','Round Upto','59030') "
VstarConn.Execute " Insert Into NewCaptions Values('60001','CO Rule No.','60001') "
VstarConn.Execute " Insert Into NewCaptions Values('60002','Give CO on','60002') "
VstarConn.Execute " Insert Into NewCaptions Values('60003','CO must be availed within','60003') "
VstarConn.Execute " Insert Into NewCaptions Values('60004','minimum for 1/2 days','60004') "
VstarConn.Execute " Insert Into NewCaptions Values('60005','Minimum for Full day','60005') "
VstarConn.Execute " Insert Into NewCaptions Values('60006','CO Rule cannot be empty','60006') "
VstarConn.Execute " Insert Into NewCaptions Values('60007','CO Rule already exists','60007') "
VstarConn.Execute " Insert Into NewCaptions Values('60008','CO Description cannot be empty','60008') "
VstarConn.Execute " Insert Into NewCaptions Values('60009','Full Day Hours should be greater than Half Day Hours.','60009') "
VstarConn.Execute " Insert Into NewCaptions Values('60010','CO Rule not found','60010') "
VstarConn.Execute " Insert Into NewCaptions Values('60011','Deduct Late Hours','60011') "
VstarConn.Execute " Insert Into NewCaptions Values('60012','Deduct Early Hours','60012') "
VstarConn.Execute " Insert Into NewCaptions Values('61001','Change User Password','61001') "
VstarConn.Execute " Insert Into NewCaptions Values('61002','Old Password','61002') "
VstarConn.Execute " Insert Into NewCaptions Values('61003','New Password','61003') "
VstarConn.Execute " Insert Into NewCaptions Values('61004','Confirm Password','61004') "
VstarConn.Execute " Insert Into NewCaptions Values('61005','Passwords don''t match','61005') "
VstarConn.Execute " Insert Into NewCaptions Values('61006','Invalid User Name','61006') "
VstarConn.Execute " Insert Into NewCaptions Values('61007','Invalid Password','61007') "
VstarConn.Execute " Insert Into NewCaptions Values('62001','OT Authorization','62001') "
VstarConn.Execute " Insert Into NewCaptions Values('62002','OT Details','62003') "
VstarConn.Execute " Insert Into NewCaptions Values('62003','Transaction File not found for the Month of','62003') "
VstarConn.Execute " Insert Into NewCaptions Values('62004','Updated OT cannot be Greater than Existing OT','62004') "
VstarConn.Execute " Insert Into NewCaptions Values('62005','OT Authorized','62005') "
VstarConn.Execute " Insert Into NewCaptions Values('62006','Work Hours','62006') "
VstarConn.Execute " Insert Into NewCaptions Values('62007','OT Hrs.','62007') "
VstarConn.Execute " Insert Into NewCaptions Values('62008','No Records Found For the Employee','62008') "
VstarConn.Execute " Insert Into NewCaptions Values('62009','Cannot Update When OT hours are 0.','') "
VstarConn.Execute " Insert Into NewCaptions Values('63001','Location Master','63001') "
VstarConn.Execute " Insert Into NewCaptions Values('63002','Location Code cannot be Blank','63002') "
VstarConn.Execute " Insert Into NewCaptions Values('63003','Location Already Exists','63003') "
VstarConn.Execute " Insert Into NewCaptions Values('63004','Description cannot be Blank','63004') "
VstarConn.Execute " Insert Into NewCaptions Values('64001','Set Employee Details','64001') "
VstarConn.Execute " Insert Into NewCaptions Values('64002','Select Details','64002') "
VstarConn.Execute " Insert Into NewCaptions Values('64003','Set &Category','64003') "
VstarConn.Execute " Insert Into NewCaptions Values('64004','Set &Department','64004') "
VstarConn.Execute " Insert Into NewCaptions Values('64005','Set &Group','64005') "
VstarConn.Execute " Insert Into NewCaptions Values('64006','Set &Location','64006') "
VstarConn.Execute " Insert Into NewCaptions Values('64007','Set &OT Rule','64007') "
VstarConn.Execute " Insert Into NewCaptions Values('64008','Set CO &Rule','64008') "
VstarConn.Execute " Insert Into NewCaptions Values('64009','Set &Entries','64009') "
VstarConn.Execute " Insert Into NewCaptions Values('64010','Set Desi&gnation','64010') "
VstarConn.Execute " Insert Into NewCaptions Values('64011','Are you sure to Change the Details ?','64011') "
VstarConn.Execute " Insert Into NewCaptions Values('64012','Set Di&vision','64012') "
VstarConn.Execute " Insert Into NewCaptions Values('65001','&Back','65001') "
VstarConn.Execute " Insert Into NewCaptions Values('65002','&Next','65002') "
VstarConn.Execute " Insert Into NewCaptions Values('65003','Add','65003') "
VstarConn.Execute " Insert Into NewCaptions Values('65004','Delete','65004') "
VstarConn.Execute " Insert Into NewCaptions Values('65005','Given Below is the List of Users currently existing in the Application. The first column depicts the name of the user while the other depicts the TYPE of the user.To View /Edit the details of any user select the USER by clicking on the GRID and click Next. Thereafter just fill in the details prompted as needed. TO ADD/DELETE users click on the respective buttons.','65005') "
VstarConn.Execute " Insert Into NewCaptions Values('65006','If you are adding a New User enter the User Name of the new user you want to add. User Name can be MAXIMUM of 20 Characters. After that select the type of user you want the user to be. There are three types of user wiz Administrator, HOD i.e Head of Department and General User. The Rest of the Process will be depending heavily on the type of user selected.','65006') "
VstarConn.Execute " Insert Into NewCaptions Values('65007','Please Select the Department for which this user will work as H.O.D. It is mandatory to select the department without which further details wont be accepted and the user record will not be saved. After that select the RIGHTS for various departmental OPERATIONS to be carried out.','65007') "
VstarConn.Execute " Insert Into NewCaptions Values('65008','Please Select the Master Rights which are to be given to this User. Master Rights are the rights which are to be given on master files such as SHIFT master, GROUP master, EMPLOYEE master etc. All of the masters can be manipulated through Adding, Editing or Deleting the Records. That is why even the rights can be alloted the same way as shown below.','65008') "
VstarConn.Execute " Insert Into NewCaptions Values('65009','Please Select the Leave Transaction Rights which are to be given to this User. Leave Transaction Rights are the rights which are to be given on Leave transactions such as Opening Leave, Credit Leave , Encash Leave and Avail Leave. Leave Transactions can be either Added or Deleted, so the rights have to be given the same way as below.','65009') "
VstarConn.Execute " Insert Into NewCaptions Values('65010','Please Select the Other Rights which are given to this User. Other Rights include rights for Shift Schedule Creation, Daily Process, Monthly Process, Correction etc.  Please Check  the boxes accordingly for the rights given below.','65010') "
VstarConn.Execute " Insert Into NewCaptions Values('65011','Please enter the Passwords (Maximum 20) for this user. If the user is added as new user, old password would not be entered, but if existing user is edited, old password will be required to  change his password. Similarly user has also to enter his second level password for CRITICAL operations like Data Correction, Leave Transaction etc.','65011') "
VstarConn.Execute " Insert Into NewCaptions Values('65012','Label for Description','65012') "

VstarConn.Execute " Insert Into NewCaptions Values('65013','Old Password','65013') "
VstarConn.Execute " Insert Into NewCaptions Values('65014','New Password','65014') "
VstarConn.Execute " Insert Into NewCaptions Values('65015','Confirm Password','65015') "
VstarConn.Execute " Insert Into NewCaptions Values('65016','Leave Rights','65016') "
VstarConn.Execute " Insert Into NewCaptions Values('65017','Master Rights','65017') "
VstarConn.Execute " Insert Into NewCaptions Values('65018','Other Rights','65018') "
VstarConn.Execute " Insert Into NewCaptions Values('65019','Install','65019') "
VstarConn.Execute " Insert Into NewCaptions Values('65020','Yearly Leave Files Rights','65020') "
VstarConn.Execute " Insert Into NewCaptions Values('65021','Shift Schedule Rights','65021') "
VstarConn.Execute " Insert Into NewCaptions Values('65022','Process Rights','65022') "
VstarConn.Execute " Insert Into NewCaptions Values('65023','Daily Data Correction Rights','65023') "
VstarConn.Execute " Insert Into NewCaptions Values('65024','Report Rights','65024') "
VstarConn.Execute " Insert Into NewCaptions Values('65025','Miscellaneous Rights','65025') "
VstarConn.Execute " Insert Into NewCaptions Values('65026','Select Rights','65026') "
VstarConn.Execute " Insert Into NewCaptions Values('65027','Opening','65027') "
VstarConn.Execute " Insert Into NewCaptions Values('65028','Credit','65028') "
VstarConn.Execute " Insert Into NewCaptions Values('65029','Encash','65029') "
VstarConn.Execute " Insert Into NewCaptions Values('65030','Avail','65030') "
VstarConn.Execute " Insert Into NewCaptions Values('65031','Transactions','65031') "
VstarConn.Execute " Insert Into NewCaptions Values('65032','If This user is assigned as ADMINISTRATOR, he will automatically be assigned all the RIGHTS and PRIVILEDGES of the application.','65032') "
VstarConn.Execute " Insert Into NewCaptions Values('65033','If This user is assigned as Manager, he will given RIGHTS for the OPERATIONS pertaining to his Selection only. The Next step will be to assign a particular Rights to this user then assign him some Selection criteria.','65033') "
VstarConn.Execute " Insert Into NewCaptions Values('65034','If This user is assigned as General User, he can be RIGHTS for all the OPERATIONS in the APPLICATION except User Management, i.e he will not be able to MANIPLUATE user accounts.','65034') "
VstarConn.Execute " Insert Into NewCaptions Values('65035','User Type','65035') "
VstarConn.Execute " Insert Into NewCaptions Values('65036','[USER NAME]','65036') "
VstarConn.Execute " Insert Into NewCaptions Values('65037','User Name','65037') "
VstarConn.Execute " Insert Into NewCaptions Values('65038','[USER TYPE]','65038') "
VstarConn.Execute " Insert Into NewCaptions Values('65039','Login Password','65039') "
VstarConn.Execute " Insert Into NewCaptions Values('65040','Second Level Password','65040') "
VstarConn.Execute " Insert Into NewCaptions Values('65041','Type of Users','65041') "
VstarConn.Execute " Insert Into NewCaptions Values('65042','List of Users','65042') "
VstarConn.Execute " Insert Into NewCaptions Values('65043','User Details','65043') "
VstarConn.Execute " Insert Into NewCaptions Values('65044','HOD Details and Rights','65044') "
VstarConn.Execute " Insert Into NewCaptions Values('65045','Master Tables Rights','65045') "
VstarConn.Execute " Insert Into NewCaptions Values('65046','Leave transaction Rights','65046') "
VstarConn.Execute " Insert Into NewCaptions Values('65047','Passwords','65047') "
VstarConn.Execute " Insert Into NewCaptions Values('65048','List of Current Users','65048') "
VstarConn.Execute " Insert Into NewCaptions Values('65049','Administrator','65049') "
VstarConn.Execute " Insert Into NewCaptions Values('65050','Manager','65050') "
VstarConn.Execute " Insert Into NewCaptions Values('65051','General User','65051') "
VstarConn.Execute " Insert Into NewCaptions Values('65052','&Edit Password','65052') "
VstarConn.Execute " Insert Into NewCaptions Values('65053','Edit Parameter','65053') "
VstarConn.Execute " Insert Into NewCaptions Values('65054','On Duty','65054') "
VstarConn.Execute " Insert Into NewCaptions Values('65055','Off Duty','65055') "
VstarConn.Execute " Insert Into NewCaptions Values('65056','OT Authorization','65056') "
VstarConn.Execute " Insert Into NewCaptions Values('65057','Edit CO','65057') "
VstarConn.Execute " Insert Into NewCaptions Values('65058','Time','65058') "
VstarConn.Execute " Insert Into NewCaptions Values('65059','General Reports','65059') "
VstarConn.Execute " Insert Into NewCaptions Values('65060','View Daily Data','65060') "
VstarConn.Execute " Insert Into NewCaptions Values('65061','Reset Locks','65061') "
VstarConn.Execute " Insert Into NewCaptions Values('65062','Export Data','65062') "
VstarConn.Execute " Insert Into NewCaptions Values('65063','Create Files','65063') "
VstarConn.Execute " Insert Into NewCaptions Values('65064','Delete Old Daily Data','65064') "
VstarConn.Execute " Insert Into NewCaptions Values('65065','Edit Paid Days','65065') "
VstarConn.Execute " Insert Into NewCaptions Values('65066','Compact Database','65066') "
VstarConn.Execute " Insert Into NewCaptions Values('65067','Backup','65067') "
VstarConn.Execute " Insert Into NewCaptions Values('65068','Restore','65068') "
VstarConn.Execute " Insert Into NewCaptions Values('65069','Update Leave Balances','65069') "
VstarConn.Execute " Insert Into NewCaptions Values('65070','Create Shift Schedule','65070') "
VstarConn.Execute " Insert Into NewCaptions Values('65071','Edit Shift Shedule','65071') "
VstarConn.Execute " Insert Into NewCaptions Values('65072','Daily Process','65072') "
VstarConn.Execute " Insert Into NewCaptions Values('65073','Monthly Process','65073') "
VstarConn.Execute " Insert Into NewCaptions Values('65074','Record','65074') "
VstarConn.Execute " Insert Into NewCaptions Values('65075','Table Name','65075') "
VstarConn.Execute " Insert Into NewCaptions Values('65076','Edit','65076') "
VstarConn.Execute " Insert Into NewCaptions Values('65077','Cannot Add New User','65077') "
VstarConn.Execute " Insert Into NewCaptions Values('65078','User Cannot Delete Himself','65078') "
VstarConn.Execute " Insert Into NewCaptions Values('65079','Please Enter the User Name','65079') "
VstarConn.Execute " Insert Into NewCaptions Values('65080','User Already Exists','65080') "
VstarConn.Execute " Insert Into NewCaptions Values('65081','Please enter LOGIN Password','65081') "
VstarConn.Execute " Insert Into NewCaptions Values('65082','LOGIN Passwords don''t match','65082') "
VstarConn.Execute " Insert Into NewCaptions Values('65083','Encryption Error::LOGIN','65083') "
VstarConn.Execute " Insert Into NewCaptions Values('65084','Please enter SECOND LEVEL Password','65084') "
VstarConn.Execute " Insert Into NewCaptions Values('65085','SECOND LEVEL Passwords don''t match','65085') "
VstarConn.Execute " Insert Into NewCaptions Values('65086','Encryption Error::SECOND LEVEL','65086') "
VstarConn.Execute " Insert Into NewCaptions Values('65087','Please Select the Department for this Manager','65087') "
VstarConn.Execute " Insert Into NewCaptions Values('65088','Are You Sure to Exit','65088') "
VstarConn.Execute " Insert Into NewCaptions Values('65089','&Select/Unselect','65089') "
VstarConn.Execute " Insert Into NewCaptions Values('65090','User Accounts Management','65090') "
VstarConn.Execute " Insert Into NewCaptions Values('65091','Please select atleast One Department for the user','65091') "
VstarConn.Execute " Insert Into NewCaptions Values('65092','Please select atleast One Company for the user','65092') "
VstarConn.Execute " Insert Into NewCaptions Values('65093','Please select atleast One Group for the user','65093') "
VstarConn.Execute " Insert Into NewCaptions Values('65094','Please select atleast One Division for the user','65094') "
VstarConn.Execute " Insert Into NewCaptions Values('65095','Please select atleast One Location for the user','65095') "
VstarConn.Execute " Insert Into NewCaptions Values('65096','Given below is the List of Masters available in the system. Please select appropriate option from each of them whichever is accessible to the HOD.','65096') "
VstarConn.Execute " Insert Into NewCaptions Values('66001','Division Master','66001') "
VstarConn.Execute " Insert Into NewCaptions Values('66002','Division not Found','66002') "
VstarConn.Execute " Insert Into NewCaptions Values('66003','Division Code cannot be blank','66003') "
VstarConn.Execute " Insert Into NewCaptions Values('66004','Division Code Already Exists','66004') "
VstarConn.Execute " Insert Into NewCaptions Values('66005','Division Name cannot be blank','66005') "
VstarConn.Execute " Insert Into NewCaptions Values('67001','    Export Data','67001') "
VstarConn.Execute " Insert Into NewCaptions Values('67002','    &Next','67002') "
VstarConn.Execute " Insert Into NewCaptions Values('67003','    &Export','67003') "
VstarConn.Execute " Insert Into NewCaptions Values('67004','    &Add','67004') "
VstarConn.Execute " Insert Into NewCaptions Values('67005','    &Remove','67005') "
VstarConn.Execute " Insert Into NewCaptions Values('67006','    A&dd All','67006') "
VstarConn.Execute " Insert Into NewCaptions Values('67007','    Rem&ove All','67007') "
VstarConn.Execute " Insert Into NewCaptions Values('67008','    &Select All','67008') "
VstarConn.Execute " Insert Into NewCaptions Values('67009','    Select &Range','67009') "
VstarConn.Execute " Insert Into NewCaptions Values('67010','    U&nselect All','67010') "
VstarConn.Execute " Insert Into NewCaptions Values('67011','    &Up','67011') "
VstarConn.Execute " Insert Into NewCaptions Values('67012','    Do&wn','67012') "
VstarConn.Execute " Insert Into NewCaptions Values('67013','    Fro&m','67013') "
VstarConn.Execute " Insert Into NewCaptions Values('67014','    T&o','67014') "
VstarConn.Execute " Insert Into NewCaptions Values('67015','    Fields Available','67015') "
VstarConn.Execute " Insert Into NewCaptions Values('67016','    Select the type of Export','67016') "
VstarConn.Execute " Insert Into NewCaptions Values('67017','    Options','67017') "
VstarConn.Execute " Insert Into NewCaptions Values('67018','    Daily Data','67018') "
VstarConn.Execute " Insert Into NewCaptions Values('67019','    Monthly Data','67019') "
VstarConn.Execute " Insert Into NewCaptions Values('67020','&Back','67020') "
VstarConn.Execute " Insert Into NewCaptions Values('67021','Please Select the Fields to be Exported','67021') "
VstarConn.Execute " Insert Into NewCaptions Values('67022','Data Exported Successfully','67022') "
VstarConn.Execute " Insert Into NewCaptions Values('67023','Some Errors Occured while saving the Export Data File','67023') "
VstarConn.Execute " Insert Into NewCaptions Values('67024','Monthly Transaction file not found for the month of','67024') "
VstarConn.Execute " Insert Into NewCaptions Values('67025','Yearly Transaction file not found for the Year of','67025') "
VstarConn.Execute " Insert Into NewCaptions Values('68001','    Change Passwords','68001') "
VstarConn.Execute " Insert Into NewCaptions Values('68002','    &Change','68002') "
VstarConn.Execute " Insert Into NewCaptions Values('68003','    Old Password','68003') "
VstarConn.Execute " Insert Into NewCaptions Values('68004','    New Password','68004') "
VstarConn.Execute " Insert Into NewCaptions Values('68005','    Confirm Password','68005') "
VstarConn.Execute " Insert Into NewCaptions Values('68006','    Login Password','68006') "
VstarConn.Execute " Insert Into NewCaptions Values('68007','    Second Level Password','68007') "
VstarConn.Execute " Insert Into NewCaptions Values('68008','    Error changing password','68008') "
VstarConn.Execute " Insert Into NewCaptions Values('68009','    Please enter old password   ','68009') "
VstarConn.Execute " Insert Into NewCaptions Values('68010','    Password retreival error:: cannot continue','68010') "
VstarConn.Execute " Insert Into NewCaptions Values('68011','    User details retreival error:: cannot continue  ','68011') "
VstarConn.Execute " Insert Into NewCaptions Values('68012','    Please enter new Password','68012') "
VstarConn.Execute " Insert Into NewCaptions Values('68013','    Passwords don''t Match','68013') "
VstarConn.Execute " Insert Into NewCaptions Values('68014','    Encryption Error::Cannot Change Password','68014') "
VstarConn.Execute " Insert Into NewCaptions Values('69001','Login','69001') "
VstarConn.Execute " Insert Into NewCaptions Values('69002','User Name','69002') "
VstarConn.Execute " Insert Into NewCaptions Values('69003','Please Enter the User Name','69003') "
VstarConn.Execute " Insert Into NewCaptions Values('69004','Invalid User Name or Password','69004') "
VstarConn.Execute " Insert Into NewCaptions Values('69005','Ambiguity in HOD''s Department of','69005') "
VstarConn.Execute " Insert Into NewCaptions Values('70001','Set Shift Details for all','70001') "
VstarConn.Execute " Insert Into NewCaptions Values('70002','Set Details For','70002') "
VstarConn.Execute " Insert Into NewCaptions Values('70003','Details of Shift Info','70003') "
VstarConn.Execute " Insert Into NewCaptions Values('70004','Details of Week Off','70004') "
VstarConn.Execute " Insert Into NewCaptions Values('70005','Details of  Additional Week Off','70005') "
VstarConn.Execute " Insert Into NewCaptions Values('70006','Details Regarding Daily Process','70006') "
VstarConn.Execute " Insert Into NewCaptions Values('70007','Details Set for Selected Employees','70007') "
VstarConn.Execute " Insert Into NewCaptions Values('D0001','Employee Code','D0001') "
VstarConn.Execute " Insert Into NewCaptions Values('D0002','Shift Code','D0002') "
VstarConn.Execute " Insert Into NewCaptions Values('D0003','Dept','D0003') "
VstarConn.Execute " Insert Into NewCaptions Values('D0004','Status','D0004') "
VstarConn.Execute " Insert Into NewCaptions Values('D0005','Employee Name','D0005') "
VstarConn.Execute " Insert Into NewCaptions Values('D0006','Remarks','D0006') "
VstarConn.Execute " Insert Into NewCaptions Values('D0007','Daily absent report for the date of','D0007') "
VstarConn.Execute " Insert Into NewCaptions Values('D0008','Page','D0008') "
VstarConn.Execute " Insert Into NewCaptions Values('D0009','Date','D0009') "
VstarConn.Execute " Insert Into NewCaptions Values('D0010','Total','D0010') "
VstarConn.Execute " Insert Into NewCaptions Values('D0011','Department :','D0011') "
VstarConn.Execute " Insert Into NewCaptions Values('D0012','Category :','D0012') "
VstarConn.Execute " Insert Into NewCaptions Values('D0013','Group :','D0013') "
VstarConn.Execute " Insert Into NewCaptions Values('D0014','Arrival Time','D0014') "
VstarConn.Execute " Insert Into NewCaptions Values('D0015','Late Hours','D0015') "
VstarConn.Execute " Insert Into NewCaptions Values('D0016','Daily arrival report for the date of','D0016') "
VstarConn.Execute " Insert Into NewCaptions Values('D0017','Late Arrival Report for the date of','D0017') "
VstarConn.Execute " Insert Into NewCaptions Values('D0018','L - Late  ::  E - Early  ::  P - With Permission','D0018') "
VstarConn.Execute " Insert Into NewCaptions Values('D0019','Absent :','D0019') "
VstarConn.Execute " Insert Into NewCaptions Values('D0020','Weekly Off :','D0020') "
VstarConn.Execute " Insert Into NewCaptions Values('D0021','Leave :','D0021') "
VstarConn.Execute " Insert Into NewCaptions Values('D0022','Present :','D0022') "
VstarConn.Execute " Insert Into NewCaptions Values('D0023','Total Late','D0023') "
VstarConn.Execute " Insert Into NewCaptions Values('D0024','Category Code','D0024') "
VstarConn.Execute " Insert Into NewCaptions Values('D0025','Name of Category','D0025') "
VstarConn.Execute " Insert Into NewCaptions Values('D0026','Late Arrival Allowed','D0026') "
VstarConn.Execute " Insert Into NewCaptions Values('D0027','Early Departure Allowed','D0027') "
VstarConn.Execute " Insert Into NewCaptions Values('D0028','Late Departure Ignore','D0028') "
VstarConn.Execute " Insert Into NewCaptions Values('D0029','Early Arrival Ignore','D0029') "
VstarConn.Execute " Insert Into NewCaptions Values('D0030','Hours Required','D0030') "
VstarConn.Execute " Insert Into NewCaptions Values('D0031','Half Day Comp.Off','D0031') "
VstarConn.Execute " Insert Into NewCaptions Values('D0032','Full Day Comp. Off','D0032') "
VstarConn.Execute " Insert Into NewCaptions Values('D0033','Category Master List','D0033') "
VstarConn.Execute " Insert Into NewCaptions Values('D0034','Continuous Absent Report for the period from','D0034') "
VstarConn.Execute " Insert Into NewCaptions Values('D0035','Emp. Code','D0035') "
VstarConn.Execute " Insert Into NewCaptions Values('D0036','Emp. Name','D0036') "
VstarConn.Execute " Insert Into NewCaptions Values('D0037','Department Code','D0037') "
VstarConn.Execute " Insert Into NewCaptions Values('D0038','Name of Department','D0038') "
VstarConn.Execute " Insert Into NewCaptions Values('D0039','Department Strength','D0039') "
VstarConn.Execute " Insert Into NewCaptions Values('D0040','Department Master List','D0040') "
VstarConn.Execute " Insert Into NewCaptions Values('D0041','Daily Early Departure report for the date of','D0041') "
VstarConn.Execute " Insert Into NewCaptions Values('D0042','Dept.Time','D0042') "
VstarConn.Execute " Insert Into NewCaptions Values('D0043','Early Hours','D0043') "
VstarConn.Execute " Insert Into NewCaptions Values('D0044','Rest Out','D0044') "
VstarConn.Execute " Insert Into NewCaptions Values('D0045','Rest In','D0045') "
VstarConn.Execute " Insert Into NewCaptions Values('D0046','Work Hours','D0046') "
VstarConn.Execute " Insert Into NewCaptions Values('D0047','Extra Out','D0047') "
VstarConn.Execute " Insert Into NewCaptions Values('D0048','Extra In','D0048') "
VstarConn.Execute " Insert Into NewCaptions Values('D0049','Daily Irregular Report','D0049') "
VstarConn.Execute " Insert Into NewCaptions Values('D0050','Daily Outdoor duty report for the date of','D0050') "
VstarConn.Execute " Insert Into NewCaptions Values('D0051','OdFrm','D0051') "
VstarConn.Execute " Insert Into NewCaptions Values('D0052','OdTo','D0052') "
VstarConn.Execute " Insert Into NewCaptions Values('D0053','OT','D0053') "
VstarConn.Execute " Insert Into NewCaptions Values('D0054','Daily Overtime Report','D0054') "
VstarConn.Execute " Insert Into NewCaptions Values('D0055','Daily Performance report for the date of','D0055') "
VstarConn.Execute " Insert Into NewCaptions Values('D0056','Summary Report for the Date Of','D0056') "
VstarConn.Execute " Insert Into NewCaptions Values('D0057','Serial No','D0057') "
VstarConn.Execute " Insert Into NewCaptions Values('D0058','Total Strength','D0058') "
VstarConn.Execute " Insert Into NewCaptions Values('D0059','No.of Emp.','D0059') "
VstarConn.Execute " Insert Into NewCaptions Values('D0060','OD','D0060') "

VstarConn.Execute " Insert Into NewCaptions Values('D0061','Employee Master Details','D0061') "
VstarConn.Execute " Insert Into NewCaptions Values('D0062','Card No','D0062') "
VstarConn.Execute " Insert Into NewCaptions Values('D0063','Designation','D0063') "
VstarConn.Execute " Insert Into NewCaptions Values('D0064','Shift Type','D0064') "
VstarConn.Execute " Insert Into NewCaptions Values('D0065','Joining Date','D0065') "
VstarConn.Execute " Insert Into NewCaptions Values('D0066','Min.entries','D0066') "
VstarConn.Execute " Insert Into NewCaptions Values('D0067','Basic Salary','D0067') "
VstarConn.Execute " Insert Into NewCaptions Values('D0068','Current Address','D0068') "
VstarConn.Execute " Insert Into NewCaptions Values('D0069','Address','D0069') "
VstarConn.Execute " Insert Into NewCaptions Values('D0070','City','D0070') "
VstarConn.Execute " Insert Into NewCaptions Values('D0071','Tel','D0071') "
VstarConn.Execute " Insert Into NewCaptions Values('D0072','PinCode','D0072') "
VstarConn.Execute " Insert Into NewCaptions Values('D0073','Sex','D0073') "
VstarConn.Execute " Insert Into NewCaptions Values('D0074','Blood group','D0074') "
VstarConn.Execute " Insert Into NewCaptions Values('D0075','Birth date','D0075') "
VstarConn.Execute " Insert Into NewCaptions Values('D0076','Permanent Address','D0076') "
VstarConn.Execute " Insert Into NewCaptions Values('D0077','District','D0077') "
VstarConn.Execute " Insert Into NewCaptions Values('D0078','State','D0078') "
VstarConn.Execute " Insert Into NewCaptions Values('D0079','PhoneNo','D0079') "
VstarConn.Execute " Insert Into NewCaptions Values('D0080','Remarks Comments','D0080') "
VstarConn.Execute " Insert Into NewCaptions Values('D0081','Left Date','D0081') "
VstarConn.Execute " Insert Into NewCaptions Values('D0082','Daily entry report for','D0082') "
VstarConn.Execute " Insert Into NewCaptions Values('D0083','Punches','D0083') "
VstarConn.Execute " Insert Into NewCaptions Values('D0084','Group Master','D0084') "
VstarConn.Execute " Insert Into NewCaptions Values('D0085','Group Code','D0085') "
VstarConn.Execute " Insert Into NewCaptions Values('D0086','Group Description','D0086') "
VstarConn.Execute " Insert Into NewCaptions Values('D0087','Holiday Master List','D0087') "
VstarConn.Execute " Insert Into NewCaptions Values('D0088','Holiday Date','D0088') "
VstarConn.Execute " Insert Into NewCaptions Values('D0089','Holiday Description','D0089') "
VstarConn.Execute " Insert Into NewCaptions Values('D0090','Leave Master List','D0090') "
VstarConn.Execute " Insert Into NewCaptions Values('D0091','Paid Leave','D0091') "
VstarConn.Execute " Insert Into NewCaptions Values('D0092','Balance','D0092') "
VstarConn.Execute " Insert Into NewCaptions Values('D0093','Encash','D0093') "
VstarConn.Execute " Insert Into NewCaptions Values('D0094','Leave days calculation','D0094') "
VstarConn.Execute " Insert Into NewCaptions Values('D0095','Yearly','D0095') "
VstarConn.Execute " Insert Into NewCaptions Values('D0096','No of Leave','D0096') "
VstarConn.Execute " Insert Into NewCaptions Values('D0097','Accumulation','D0097') "
VstarConn.Execute " Insert Into NewCaptions Values('D0098','Employee Master List','D0098') "
VstarConn.Execute " Insert Into NewCaptions Values('D0099','Daily Manpower report for the date of','D0099') "
VstarConn.Execute " Insert Into NewCaptions Values('D0100','Present','D0100') "
VstarConn.Execute " Insert Into NewCaptions Values('D0101','Absent','D0101') "
VstarConn.Execute " Insert Into NewCaptions Values('D0102','Offs','D0102') "
VstarConn.Execute " Insert Into NewCaptions Values('D0103','Monthly Performance report for the month of','D0103') "
VstarConn.Execute " Insert Into NewCaptions Values('D0104','Overtime report for the month of','D0104') "
VstarConn.Execute " Insert Into NewCaptions Values('D0105','L A T E S','D0105') "
VstarConn.Execute " Insert Into NewCaptions Values('D0106','E A R L Y','D0106') "
VstarConn.Execute " Insert Into NewCaptions Values('D0107','No.','D0107') "
VstarConn.Execute " Insert Into NewCaptions Values('D0108','Absent / Late / Early Summary report for the month of','D0108') "
VstarConn.Execute " Insert Into NewCaptions Values('D0109','Monthly attendance for the month of','D0109') "
VstarConn.Execute " Insert Into NewCaptions Values('D0110','OT_HRS','D0110') "
VstarConn.Execute " Insert Into NewCaptions Values('D0111','Total Early','D0111') "
VstarConn.Execute " Insert Into NewCaptions Values('D0112','Monthly Late Arrival Report for the month of','D0112') "
VstarConn.Execute " Insert Into NewCaptions Values('D0113','Monthly Early Departure Report for the month of','D0113') "
VstarConn.Execute " Insert Into NewCaptions Values('D0114','Leave Balance report for the month of','D0114') "
VstarConn.Execute " Insert Into NewCaptions Values('D0115','Leave Code','D0115') "
VstarConn.Execute " Insert Into NewCaptions Values('D0116','Leave Name','D0116') "
VstarConn.Execute " Insert Into NewCaptions Values('D0117','Leave From','D0117') "
VstarConn.Execute " Insert Into NewCaptions Values('D0118','Leave To','D0118') "
VstarConn.Execute " Insert Into NewCaptions Values('D0119','Leave Days','D0119') "
VstarConn.Execute " Insert Into NewCaptions Values('D0120','Time','D0120') "
VstarConn.Execute " Insert Into NewCaptions Values('D0121','2nd Punch','D0121') "
VstarConn.Execute " Insert Into NewCaptions Values('D0122','3rd Punch','D0122') "
VstarConn.Execute " Insert Into NewCaptions Values('D0123','4th Punch','D0123') "
VstarConn.Execute " Insert Into NewCaptions Values('D0124','5th Punch','D0124') "
VstarConn.Execute " Insert Into NewCaptions Values('D0125','Leave Availed for the Month of','D0125') "
VstarConn.Execute " Insert Into NewCaptions Values('D0190','Early Departure memo for the month of','D0190') "
VstarConn.Execute " Insert Into NewCaptions Values('D0191','Late Arrival Memo for the month of','D0191') "
VstarConn.Execute " Insert Into NewCaptions Values('D0192','Absent Memo for the month of','D0192') "
VstarConn.Execute " Insert Into NewCaptions Values('D0193','O.T. Paid','D0193') "
VstarConn.Execute " Insert Into NewCaptions Values('D0194','O.T. Done','D0194') "
VstarConn.Execute " Insert Into NewCaptions Values('D0195','Overtime Paid report for the month of','D0195') "
VstarConn.Execute " Insert Into NewCaptions Values('D0196','Monthly Absent  Report for the month of','D0196') "
VstarConn.Execute " Insert Into NewCaptions Values('D0197','Monthly Present Report for the month of','D0197') "
VstarConn.Execute " Insert Into NewCaptions Values('D0198','Monthly muster for the month of','D0198') "
VstarConn.Execute " Insert Into NewCaptions Values('D0199','Monthly shift schedule report for the month of','D0199') "
VstarConn.Execute " Insert Into NewCaptions Values('D0201','Yearly Performance Report for the year','D0201') "
VstarConn.Execute " Insert Into NewCaptions Values('D0202','PAIDDAYS','D0202') "
VstarConn.Execute " Insert Into NewCaptions Values('D0203','LT_NO','D0203') "
VstarConn.Execute " Insert Into NewCaptions Values('D0204','LT_HRS','D0204') "
VstarConn.Execute " Insert Into NewCaptions Values('D0205','ERL_NO','D0205') "
VstarConn.Execute " Insert Into NewCaptions Values('D0206','ERL_HRS','D0206') "
VstarConn.Execute " Insert Into NewCaptions Values('D0207','WRK_HRS','D0207') "
VstarConn.Execute " Insert Into NewCaptions Values('D0208','Yearly Man-Power Utilisation Report for the year','D0208') "
VstarConn.Execute " Insert Into NewCaptions Values('D0209','NIGHT','D0209') "
VstarConn.Execute " Insert Into NewCaptions Values('D0210','Leave Utilisation Report for year','D0210') "
VstarConn.Execute " Insert Into NewCaptions Values('D0211','Leave Code','D0211') "
VstarConn.Execute " Insert Into NewCaptions Values('D0212','Remarks / Period','D0212') "
VstarConn.Execute " Insert Into NewCaptions Values('D0213','Credited','D0213') "
VstarConn.Execute " Insert Into NewCaptions Values('D0214','Availed','D0214') "
VstarConn.Execute " Insert Into NewCaptions Values('D0215','Yearly Absent Report for the year','D0215') "
VstarConn.Execute " Insert Into NewCaptions Values('D0216','Yearly Present Report for the year','D0216') "
VstarConn.Execute " Insert Into NewCaptions Values('D0217','Weekly Performance report for the period of','D0217') "
VstarConn.Execute " Insert Into NewCaptions Values('D0218','To','D0218') "
VstarConn.Execute " Insert Into NewCaptions Values('D0219','Arr','D0219') "
VstarConn.Execute " Insert Into NewCaptions Values('D0220','Dep','D0220') "
VstarConn.Execute " Insert Into NewCaptions Values('D0221','Late','D0221') "
VstarConn.Execute " Insert Into NewCaptions Values('D0222','Earl','D0222') "
VstarConn.Execute " Insert Into NewCaptions Values('D0223','Work','D0223') "
VstarConn.Execute " Insert Into NewCaptions Values('D0224','Rem','D0224') "
VstarConn.Execute " Insert Into NewCaptions Values('D0225','Shf','D0225') "
VstarConn.Execute " Insert Into NewCaptions Values('D0226','Weekly Absent report for the period of','D0226') "
VstarConn.Execute " Insert Into NewCaptions Values('D0227','Weekly Attendance report for the period of','D0227') "
VstarConn.Execute " Insert Into NewCaptions Values('D0228','Weekly Early Departure report for the period of','D0228') "
VstarConn.Execute " Insert Into NewCaptions Values('D0229','Weekly Late Arrival report for the period of','D0229') "
VstarConn.Execute " Insert Into NewCaptions Values('D0230','Weekly Overtime report for the period of','D0230') "
VstarConn.Execute " Insert Into NewCaptions Values('D0231','Weekly Shift Arrangement report for the period of','D0231') "
VstarConn.Execute " Insert Into NewCaptions Values('D0232','Weekly Irregular Punch report from','D0232') "
VstarConn.Execute " Insert Into NewCaptions Values('D0233','Shift Master List','D0233') "
VstarConn.Execute " Insert Into NewCaptions Values('D0234','Name of the shift','D0234') "
VstarConn.Execute " Insert Into NewCaptions Values('D0235','Shift Time','D0235') "
VstarConn.Execute " Insert Into NewCaptions Values('D0236','Starting','D0236') "
VstarConn.Execute " Insert Into NewCaptions Values('D0237','Ending','D0237') "
VstarConn.Execute " Insert Into NewCaptions Values('D0238','Hours','D0238') "
VstarConn.Execute " Insert Into NewCaptions Values('D0239','Lunch Time','D0239') "
VstarConn.Execute " Insert Into NewCaptions Values('D0240','Half Day Time','D0240') "
VstarConn.Execute " Insert Into NewCaptions Values('D0241','Daily Shift arrangement report for the date of','D0241') "
VstarConn.Execute " Insert Into NewCaptions Values('D0242','Rotational Shift Master List','D0242') "
VstarConn.Execute " Insert Into NewCaptions Values('D0243','Rotation Code','D0243') "
VstarConn.Execute " Insert Into NewCaptions Values('D0244','Type of Rotation','D0244') "
VstarConn.Execute " Insert Into NewCaptions Values('D0245','Shift Pattern','D0245') "
VstarConn.Execute " Insert Into NewCaptions Values('D0246','Rotation Pattern','D0246') "
VstarConn.Execute " Insert Into NewCaptions Values('D0247','Overtime Report for the Period from','D0247') "
VstarConn.Execute " Insert Into NewCaptions Values('D0248','Performance report from','D0248') "
VstarConn.Execute " Insert Into NewCaptions Values('D0249','Attendance Muster report for the period from','D0249') "
VstarConn.Execute " Insert Into NewCaptions Values('D0250','Late Arrival report for the period from','D0250') "
VstarConn.Execute " Insert Into NewCaptions Values('D0251','Early Departure report for the period from','D0251') "
VstarConn.Execute " Insert Into NewCaptions Values('D0252','WO on holiday report for the month of','D0252') "
VstarConn.Execute " Insert Into NewCaptions Values('D0253','Days to be deducted','D0253') "
VstarConn.Execute " Insert Into NewCaptions Values('D0254','Total Lates report for the month of','D0254') "
VstarConn.Execute " Insert Into NewCaptions Values('D0255','Total Early report for the month of','D0255') "
VstarConn.Execute " Insert Into NewCaptions Values('D0256','Division Master','D0256') "
VstarConn.Execute " Insert Into NewCaptions Values('D0257','Strength','D0257') "
VstarConn.Execute " Insert Into NewCaptions Values('D0258','M 1','D0258') "
VstarConn.Execute " Insert Into NewCaptions Values('D0259','M 2','D0259') "
VstarConn.Execute " Insert Into NewCaptions Values('D0260','Meal Allowance Report From','D0260') "
VstarConn.Execute " Insert Into NewCaptions Values('D0261','Summary report from','D0261') "
VstarConn.Execute " Insert Into NewCaptions Values('D0262','Permission card report From','D0262') "
VstarConn.Execute " Insert Into NewCaptions Values('D0263','Leave Availment Report From ','D0263') "
VstarConn.Execute " Insert Into NewCaptions Values('D0264','Corrected','D0264') "
VstarConn.Execute " Insert Into NewCaptions Values('D0265 ','C','D0265') "
VstarConn.Execute " Insert Into NewCaptions Values('M1001','Invalid Date Format Found:: Cannot Proceed','M1001') "
VstarConn.Execute " Insert Into NewCaptions Values('M1002','Please set Your Date Settings to ''M/D/YY'' Format in','M1002') "
VstarConn.Execute " Insert Into NewCaptions Values('M1003','Start --> Settings --> Control Panel --> Regional Settings --> Date','M1003') "
VstarConn.Execute " Insert Into NewCaptions Values('M1004','--> Short Date Style','M1004') "
VstarConn.Execute " Insert Into NewCaptions Values('M1005','Please set Your Date Settings to ''DD/MM/YY'' Format in','M1005') "
VstarConn.Execute " Insert Into NewCaptions Values('M1006','Your Regional Date Settings do not Match the Application Date Settings','M1006') "
VstarConn.Execute " Insert Into NewCaptions Values('M1007','(American Type)','M1007') "
VstarConn.Execute " Insert Into NewCaptions Values('M1008','Do You Wish to Set It','M1008') "
VstarConn.Execute " Insert Into NewCaptions Values('M1009','( British Type)','M1009') "
VstarConn.Execute " Insert Into NewCaptions Values('M1010','One or More Required Script Files are Missing :: Cannot Run Daily Process.','M1010') "
VstarConn.Execute " Insert Into NewCaptions Values('M1011','Total Calculation File Needed for Monthly Processing not Found','M1011') "
VstarConn.Execute " Insert Into NewCaptions Values('M1012','Rules File Needed for Monthly Processing not Found','M1012') "
VstarConn.Execute " Insert Into NewCaptions Values('M1013','Error Loading Total Calculation File Needed for Monthly Processing','M1013') "
VstarConn.Execute " Insert Into NewCaptions Values('M1014','Error Loading Rules File Needed for Monthly Processing','M1014') "
VstarConn.Execute " Insert Into NewCaptions Values('M1015','[Demo Version]','M1015') "
VstarConn.Execute " Insert Into NewCaptions Values('M3001','CO not Found :: Leave Balance File for the Current Year not Updated','M3001') "
VstarConn.Execute " Insert Into NewCaptions Values('M3002','CO not Found in Leave Master','M3002') "
VstarConn.Execute " Insert Into NewCaptions Values('M3003','Leave Balanace File for the Current Year not Found','M3003') "
VstarConn.Execute " Insert Into NewCaptions Values('M3004','Please Create it First and then do the Daily Process','M3004') "
VstarConn.Execute " Insert Into NewCaptions Values('M3005','Reading from File','M3005') "
VstarConn.Execute " Insert Into NewCaptions Values('M4001','Month should be between 01 to 12.','M4001') "
VstarConn.Execute " Insert Into NewCaptions Values('M4002','Invalid Number of Days','M4002') "
VstarConn.Execute " Insert Into NewCaptions Values('M4003','Invalid Date Length','M4003') "
VstarConn.Execute " Insert Into NewCaptions Values('M4004','Invalid Date Structure','M4004') "
VstarConn.Execute " Insert Into NewCaptions Values('M6001','Yearly Leave Files are not Created :: Please Create them','M6001') "
VstarConn.Execute " Insert Into NewCaptions Values('M6002','Yearly Leave Files for the Year','M6002') "
VstarConn.Execute " Insert Into NewCaptions Values('M6003','Yearly Updation Aborted','M6003') "
VstarConn.Execute " Insert Into NewCaptions Values('M6004','Yearly Leaves Files Updated','M6004') "
VstarConn.Execute " Insert Into NewCaptions Values('M6005','Yearly Files are not Created Properly :: Please Re-Create it.','M6005') "
VstarConn.Execute " Insert Into NewCaptions Values('M7001','Shift File for the Month Of','M7001') "
VstarConn.Execute " Insert Into NewCaptions Values('M7002','Please Select the required Report.','M7002') "
VstarConn.Execute " Insert Into NewCaptions Values('M7003','Invalid File Name','M7003') "
VstarConn.Execute " Insert Into NewCaptions Values('M7004',' saved successfully','M7004') "
VstarConn.Execute " Insert Into NewCaptions Values('M7005','Yearly Leave Transaction File Not Found','M7005') "
VstarConn.Execute " Insert Into NewCaptions Values('M7006','Mail','M7006') "
VstarConn.Execute " Insert Into NewCaptions Values('M7007','View','M7007') "
VstarConn.Execute " Insert Into NewCaptions Values('M7008','Print','M7008') "
VstarConn.Execute " Insert Into NewCaptions Values('M7009','Print to File','M7009') "
VstarConn.Execute " Insert Into NewCaptions Values('M7010','   Checking Validations ..','M7010') "
VstarConn.Execute " Insert Into NewCaptions Values('M7011','   Processing Valid data ..','M7011') "
VstarConn.Execute " Insert Into NewCaptions Values('M7012','   Executing Query ..','M7012') "
VstarConn.Execute " Insert Into NewCaptions Values('M7013','   Operation Aborted','M7013') "
VstarConn.Execute " Insert Into NewCaptions Values('M7014','   Preparing Report to','M7014') "
VstarConn.Execute " Insert Into NewCaptions Values('M7015','   Daily Reports','M7015') "
VstarConn.Execute " Insert Into NewCaptions Values('M7016','   Weekly Reports','M7016') "
VstarConn.Execute " Insert Into NewCaptions Values('M7017','   Monthly Reports','M7017') "
VstarConn.Execute " Insert Into NewCaptions Values('M7018','   Yearly Reports','M7018') "
VstarConn.Execute " Insert Into NewCaptions Values('M7019','   Masters Reports','M7019') "
VstarConn.Execute " Insert Into NewCaptions Values('M7020','   Periodic Reports','M7020') "
VstarConn.Execute " Insert Into NewCaptions Values('M7021','    Sending Mail to','M7021') "
VstarConn.Execute " Insert Into NewCaptions Values('M7022','Opening','M7022') "
VstarConn.Execute " Insert Into NewCaptions Values('M7023','Credited','M7023') "
VstarConn.Execute " Insert Into NewCaptions Values('M7024','Encashed','M7024') "
VstarConn.Execute " Insert Into NewCaptions Values('M7025','Late Cut','M7025') "
VstarConn.Execute " Insert Into NewCaptions Values('M7026','Early Cut','M7026') "
InsertSQLCaptions = True
Exit Function
Err_P:
    MsgBox Err.Description
    Resume Next
End Function
'' apoorva
Public Sub InsertSQLData()
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute " Insert Into catdesc (CAT, [DESC], invisible) Values('100','Print','Y')"
VstarConn.Execute " Insert Into CORul(COCode) Values('100')"
VstarConn.Execute " Insert Into OTRul(OTCode) Values('100')"
VstarConn.Execute " Insert Into instshft(Shift) Values('100')"
VstarConn.Execute " Insert Into Ro_shift(Scode) Values('100')"
End Sub
'' apoorva
Public Sub InsertoracleData()
VstarConn.Execute "Commit"
VstarConn.Execute " Insert Into catdesc (CAT, ""Desc"", INVISIBLE) Values('100','Print','Y')"
VstarConn.Execute " Insert Into CORul(COCode) Values('100')"
VstarConn.Execute " Insert Into OTRul(OTCode) Values('100')"
VstarConn.Execute " Insert Into instshft(Shift) Values('100')"
VstarConn.Execute " Insert Into Ro_shift(Scode) Values('100')"
VstarConn.Execute "Commit"
End Sub

'' apoorva

Public Sub InsertSQLIndex()
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute "Use VSTARDB"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('CATDESC','UNIQUE','CAT, DYS','CATDESC_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DEPTDESC','UNIQUE','DEPT','DEPTDESC_DEPT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('EMPMST','UNIQUE','EMPCODE','EMPMST_EMPCODE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('EMPMST','UNIQUE','CARD','EMPMST_CARD')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('EMPMST','UNIQUE','EMPCODE,CAT','EMPMST_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('HOLIDAY','UNIQUE','DATE, CAT','HOLIDAY_DATE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('INSTSHFT','UNIQUE','SHIFT','INSTSHFT_SHIFT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVDESC','UNIQUE','LVCODE, CAT','LEAVDESC_CODE_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVDESC','','CAT','LEAVDESC_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVDESC','','LEAVE','LEAVDESC_LEAVE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LOST','','EMPCODE, DATE, T_PUNCH','LOST')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('RO_SHIFT','UNIQUE','SCODE','RO_SHIFT_SCODE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('OTRULE','UNIQUE','OTCODE','OTRULE_OTCODE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DAILYPRO','','EMPCODE, DTE, T_PUNCH','DAILYPRO')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DTRN','UNIQUE','EMPCODE, MNDATE','DTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DIVISION','UNIQUE','DIV','DIVISION_DIV')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LATEERL','UNIQUE','EMPCODE, DATE','LATEERL')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LOCATION','UNIQUE','LOCATION','LOCATION')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('USERACCS','UNIQUE','USERNAME','USERACCS_USERNAME')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVTRN','UNIQUE','EMPCODE, LST_DATE','LVTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LVTRNYY','UNIQUE','EMPCODE, LST_DATE','LVTRNYY')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVBAL','UNIQUE','EMPCODE','LVBAL')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LVBALYY','UNIQUE','EMPCODE','LVBALYY')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVINFO','','EMPCODE, FROMDATE','LVINFO')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LVINFOYY','','EMPCODE, FROMDATE','LVINFOYY')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('MONTRN','UNIQUE','EMPCODE, DATE','MONTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('MONYYTRN','UNIQUE','EMPCODE, DATE','MONYYTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('SHFINFO','UNIQUE','EMPCODE','SHFINFO')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('MONYYSHF','UNIQUE','EMPCODE','MONYYSHF')"

End Sub
'' apoorva
Public Sub InsertOracleIndex()
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('CATDESC','UNIQUE','CAT, DYS','CATDESC_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DEPTDESC','UNIQUE','DEPT','DEPTDESC_DEPT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('EMPMST','UNIQUE','EMPCODE','EMPMST_EMPCODE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('EMPMST','UNIQUE','CARD','EMPMST_CARD')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('EMPMST','UNIQUE','EMPCODE,CAT','EMPMST_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('HOLIDAY','UNIQUE','""DATE"", CAT','HOLIDAY_DATE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('INSTSHFT','UNIQUE','SHIFT','INSTSHFT_SHIFT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVDESC','UNIQUE','LVCODE, CAT','LEAVDESC_CODE_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVDESC','','CAT','LEAVDESC_CAT')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVDESC','','LEAVE','LEAVDESC_LEAVE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LOST','','EMPCODE, ""DATE"", T_PUNCH','LOST')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('RO_SHIFT','UNIQUE','SCODE','RO_SHIFT_SCODE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('OTRULE','UNIQUE','OTCODE','OTRULE_OTCODE')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DAILYPRO','','EMPCODE, DTE, T_PUNCH','DAILYPRO')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DTRN','UNIQUE','EMPCODE, MNDATE','DTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('DIVISION','UNIQUE','DIV','DIVISION_DIV')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LATEERL','UNIQUE','EMPCODE, ""DATE""','LATEERL')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LOCATION','UNIQUE','LOCATION','LOCATION')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('USERACCS','UNIQUE','USERNAME','USERACCS_USERNAME')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVTRN','UNIQUE','EMPCODE, LST_DATE','LVTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LVTRNYY','UNIQUE','EMPCODE, LST_DATE','LVTRNYY')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVBAL','UNIQUE','EMPCODE','LVBAL')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LVBALYY','UNIQUE','EMPCODE','LVBALYY')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LEAVINFO','','EMPCODE, FROMDATE','LVINFO')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('LVINFOYY','','EMPCODE, FROMDATE','LVINFOYY')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('MONTRN','UNIQUE','EMPCODE, ""DATE""','MONTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('MONYYTRN','UNIQUE','EMPCODE, ""DATE""','MONYYTRN')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('SHFINFO','UNIQUE','EMPCODE','SHFINFO')"
VstarConn.Execute " INSERT INTO TBLINDEXMASTER VALUES ('MONYYSHF','UNIQUE','EMPCODE','MONYYSHF')"
VstarConn.Execute "Commit"
End Sub
