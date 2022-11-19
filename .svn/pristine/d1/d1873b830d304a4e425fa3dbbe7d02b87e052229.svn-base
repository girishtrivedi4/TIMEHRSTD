Attribute VB_Name = "DJModule"
Option Explicit
'' Constants
'' For Activity Log
Public Const LOCALE_SYSTEM_DEFAULT = &H400
Public Const LOCALE_SSHORTDATE As Long = &H1F
Public Const lg_NoModeAction = "GEN"
Public Const lgADD_MODE = "ADD"
Public Const lgEdit_Mode = "EDT"
Public Const lgDelete_Action = "DEL"
Public Const lgPeriodic_Action = "PER"
Public Const lgMaster_Action = "MAS"
Public Const lgReset_Action = "RES"
Public Const lgShift_MODE = "SHF"
Public Const lgRecord_MODE = "REC"
Public Const lgStatus_MODE = "STA"
Public Const lgOnDuty_MODE = "OND"
Public Const lgOffDuty_MODE = "OFD"
Public Const lgOT_MODE = "OVT"
Public Const lgTime_MODE = "TIM"
Public Const lgExclusive_Action = "EXC"
Public Const lgDaily_Action = "DLY"
Public Const lgMonthly_Action = "MON"
Public Const lgYearly_Action = "YLY"
Public Const lgLeaves_Mode = "LVS"
Public Const lgCorrection_Mode = "COR"
Public Const lgBackUp_Mode = "BAK"
Public Const lgRestore_Mode = "RES"
Public Const lgRDaily_Action = "DLY"
Public Const lgRMonthly_Action = "MON"
Public Const lgRReports_Action = "REP"
'' For API's
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
''Public Const KEY_ALL_ACCESS As Long = &HF0063   '' Previously Used Constant
Public Const KEY_ALL_ACCESS As Long = &H3F
Public Const ERROR_SUCCESS As Long = 0
Public Const WM_SETTINGCHANGE = &H1A
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_WININICHANGE = &H1A
'' Criteria Constant
Public Const SELCRIT = "DEPT"   '-- Supriya FairField 17/02/05 to remove standard error
Public Const SELCRIT1 = "LOCATION"
Public Const SELCRIT2 = "COMPANY"
Public Const SELCRIT3 = "DIV"

'' Color Constants
Public Const SELECTED_COLOR = &HC0FFFF
Public Const UNSELECTED_COLOR = vbWhite
'' Print User
Public Const strPrintUser = "TIMEHR"
Public Const strPrintPass = "RHEMIT"
'' Constant for the Number of Master Tables
Public Const TOTAL_MASTER_TABLES = 13
Public Const TOTAL_LEAVE_RIGHTS = 8
Public Const TOTAL_OTHER_RIGHTS = 20
'' Constants for YES and NO
Public Const CON_YES = "YES"
Public Const CON_NO = "NO"
'' Constants for UserType
Public Const ADMIN = "ADMIN"
Public Const HOD = "HEAD"
Public Const GENERAL = "GENERAL"
'******
Public deptsel, catsel, cboGroup, cboLoc As String
Public Const Dbname = "Attendodb"  '"VStarDB"
'Public Const Dbname = "Zuari"
Public LoginStatus As Boolean
Public Hardboot As String
'' API 's
'' 01. GetDiskFreeSpaceEx
Public Declare Function GetDiskFreeSpaceEx Lib "KERNEL32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpBytesAvailable As Currency, lpTotalBytes As Currency, lpFreeBytes As Currency) As Long
'' 02. Sleep
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
'' 03. RegCreateKey
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'' 04. RegSetValueEx
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'' 05. RegCloseKey
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'' 06. GetSystemDirectory
Public Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'' 07. RegOpenKeyEx
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'' 08. RegQueryValueEx
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
'' 09. SetLocaleInfo
Public Declare Function SetLocaleInfo Lib "KERNEL32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
'' 10. GetLocaleInfo
Public Declare Function GetLocaleInfo Lib "KERNEL32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'' 11. GetSystemDefaultLCID
Public Declare Function GetSystemDefaultLCID Lib "KERNEL32" () As Long
'' 12. PostMessage
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'' 13. FreeConsole
Public Declare Function FreeConsole Lib "KERNEL32" () As Long
''14 closehandle
'Public Declare Function closehandle Lib "kernel32" () As Long
'' Types
'' 01. ShiftInfo
Public Type ShiftInfo
        Empcode As String
        ShiftType As String
        ShiftCode As String
        startdate As Date
        WO As String
        WO1 As String
        WO2 As String
        WO3  As String
        'WO4 add by  for TGL
        WO4 As String
        ''
        'WORule Add by  For Robosoft
        WORule As String
        ''
        WOHLAction As Byte          '' 0=Previous Day Shift, 1=Next Day Shift
                                    '' 2=Specific Shift, 3= Assign the following shift
        Action3Shift As String      '' Specific Shift
        AutoOnPunch As Boolean      '' Auto Shift on Punch found
        ActionBlank As String       '' Blank Shift or Else
End Type
Public Shft As ShiftInfo
''02. Ovstar
Public Type Ovstar
        CodeSize As Integer
        CardSize As Integer
        Yearstart As String
        YearSel As String
        WeekStart As String
        Use_Mail As Boolean
        Cust_code As Boolean
        PrsCode As String
        AbsCode As String
        HlsCode  As String
        WosCode As String
        late As String
        Empcode As String
End Type
Public pVStar As Ovstar
'' 03. IniDat
Public Type IniDat
        blnVerType As String
        blnNetType As String
        blnAssum As String
        lngEmp As String
        bytUse As String
        bytCom As String
        strLok As String
        strCOM As String
        strBak As String
        strSer As String
        strVer As String
        blnWeb As String
        strUser As String
        strPass As String
        strLoc As String
End Type
Public InVar As IniDat
'' 04. MnlPr
Public Type MnlPr
    strFrtDate As String
    strLstDate As String
    strLvtDate As String
    strEmpList As String
    bytExeLE As Byte
    bytExeLunchLt As Byte
End Type
Public typMnlVar As MnlPr
'' 06. DMY
Public Type DMY
    bytD As Byte
    bytM As Byte
    BytY As Integer
End Type
Public typDMY As DMY
'' 07. LogDet
Public Type LogDet
    sngTime As Single
    strDate As String
    strUsername As String
    lngRecord As Long
    strModeType As String
    bytTranType As Byte
    bytTranSource As Byte
End Type
Public typLog As LogDet

Public Type AuditLog    ' 29-05
    sngTmt As String
    strDt As String
    strUser As String
    strMsg As String
    lngRecNum As Long
    strIn As String
    strout As String
    strIp As String
    dte As String
End Type
Public typAudit As AuditLog

'' 08. LvDesc
Public Type LvDesc
    strLvCode As String * 2     '' Leave Code
    strCat As String * 3        '' Category
    strLvName As String         '' Leave Description
    sngQty As Single            '' Leave Quantity
    sngAccQty As Single         '' Leave Accumulation Quantity
    blnCarry As Boolean         '' Carry Balance Forward
    blnCrImd As Boolean         '' Credit Immidiately or Next Year
    blnFullPro As Boolean       '' Credit Full or Proportionately
    blnLvType As Boolean        '' Leave Type
End Type
Public typLvD As LvDesc
'' 09. LvInfo
Public Type LvInfo
    strFrom As String * 10      '' From Date
    strTo As String * 10        '' To Date
    strEntry As String * 10     '' Entry Date
    sngQty As Single            '' Leave Quantity
End Type
Public typLvI As LvInfo
'' 10. RepVars
Public Type RepVars
    strDlyDate As String
    strWkDate As String
    strMonMth As String
    strMonYear As String
    strYear As String
    strLeftFr As String
    strLeftTo As String
    strPeriFr As String
    strPeriTo As String
End Type
Public typRep As RepVars
'' 11. DSNDetails
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
'' 12. RepOpts
Public Type RepOpts
    bytDly As Byte
    bytWek As Byte
    bytMon As Byte
    bytYer As Byte
    bytMst As Byte
    bytPer As Byte
End Type
Public typOptIdx As RepOpts

'
Public Type DailyLvBal
    bytMnthOpt As Byte
    bytDtOpt As Byte
    DailyDt As Date
    strMnth As String
    strYr As String
    typFdate As String
    typLdate As String
End Type
Public typDlyLvBal As DailyLvBal

Public Type MaleFemale            '
maleopt As Byte
femaleopt As Byte
End Type
Public typMaleFemale As MaleFemale
'' 13. StartEndNums
Public Type StartEndNums
    bytStart As Byte        '' Start Number
    bytEnd As Byte          '' End Number
End Type
Public typSENum As StartEndNums
'' 14. PointArrays
Public Type PointArrays
    bytWeek As Byte         '' Pointer to Week Pattern Array
    bytShift As Byte        '' Pointer to Shift Array
    bytSND As Byte          '' Pointer to Specific Days Pattern Array
    bytFD As Byte           '' Pinter to the Fixed days Array
End Type
Public typPA As PointArrays
'' 15. EmpDetRot
Public Type EmpDetRot
    strCode As String           '' Employee Code
    strName As String           '' Employee Name
    strCat As String            '' Employee Category
    strOff As String            '' First Week Off
    strOff2 As String           '' Second Week Off
    strOff_1_3 As String        '' First and Third Week Off
    strOff_2_4 As String        '' Second and Fourth Week Off
    strShifttype As String      '' Type of Shift i.e Fixed or Rotation
    strShiftCode As String      '' Shift Code e.g. ('G' for Fixed OR 'R1' for Rotation)
    dtShift As Date             '' For Shift Date
    dtJoin As Date              '' For Joining Date
    dtLeave As Date             '' For LeaveDate
    strLocation As String       '' Location     ' 18-01
End Type
Public typEmpRot As EmpDetRot

'' Variables
'' Intrinsic Variables
Public varCalDt As Variant          '' DateCalendar
Public blnDType As Boolean          '' Boolean DataSource Type
Public blnBackRes As Boolean        '' Boolean BackUp & Restore
Public strDjFileN As String         '' FileName

Public strName As String            ''for Name keyword
Public strDTEnc As String           '' Date Enclosure
Public bytMode As Byte              '' Byte Variable to Keep Track of the Modes
Public bytBackEnd                   '' Byte Variable to KeepTrack of the BackEnd
Public bytShfMode As Byte           '' Mode to Swith Between Shift Forms.
Public bytDateF As Byte             '' Byte Variable to Keep Track of DateFormat
Public strDateFO As String          '' String Variable to Keep the Date Format
Public strBackEndPath As String     '' For the BackEndPath if Access
Public bytFormToLoad As Byte        '' For the Type of Form to be Loaded
Public strRotPass As String         '' For Passing of Variables between Rots.
Public strCapSND As String          '' For Passing of Caption to the frmRotSND Form
Public blnScripts As Boolean        '' Scripts Load
Public blnDiff As Boolean           '' Different User
Public UserName As String           '' User Name
Public Msr_no As Integer
Public i As Integer                 '' Temporary Variable
Public EmailSendOpt                 '' Email Sending Option
Public EmpId As String              '' Employee Code for Emailing Reports
Public EmailSub As String           '' Email Subject
Public EmailSend As Boolean         '' Check for Email Send Option
Public AddRights As Boolean         '' For Add Rights
Public EditRights As Boolean        '' For Edit Rights
Public DeleteRights As Boolean      '' For Delete Rights
Public intUserNum As Integer        '' Login User Number
Public strCapField As String        '' Caption Field Name String
Public strFormCommand As String     '' Form Number for Captions
Public blnDBCompacted As Boolean    '' Check id Database is Compacted
Public blnShowLang As Boolean       '' For Language Features
Public blnIPAddress As Boolean
Public bytLstInd As Byte            '' Byte for Exchange of ListIndex values
Public strCurrentUserType As String '' For the Current User
''For Mauritius 09-8-2003
''Original-> Public intCurrDept As Integer       '' Current Dept of HOD who has logged in
Public strCurrDept As String       '' Current Dept of HOD who has logged in
Public strCurrData As String
''
Public strHODRights As String       '' Rights of HOD User Currently Logged in
Public strMasterRights As String    '' Master Rights of User Currently Logged in
Public strLeaveRights As String     '' Leave Rights of User Currently Logged in
Public strOtherRights1 As String    '' Other Rights of User Currently Logged in
Public strPassword As String        '' User Login password of Current User
Public strOtherPass1 As String      '' Second Level Password of Current User
'' Keywords
Public strKDate As String           '' Date keyword
Public strKGroup As String          '' Group
Public strKDesc As String           '' Desc
Public strKOff As String            '' Off
Public StrKConcat  As String         '' Concatination operator
Public CatFlag As Boolean

Public strLeft As String        ''for left
Public strRight As String       ''for right
'Public DosCatFlag As Boolean
''Database Password
Public strDBPass As String
'' Database Recordsets
Public adRsInstall As New ADODB.Recordset
Public adrsTemp As New ADODB.Recordset
Public adrstemp1 As New ADODB.Recordset
Public adrstemp2 As New ADODB.Recordset
Public adrstemp3 As New ADODB.Recordset
Public adRsintshft As New ADODB.Recordset
Public AdrsCat As New ADODB.Recordset
Public adrsLeave As New ADODB.Recordset
Public adrsEmp As New ADODB.Recordset
Public adrsemp1 As New ADODB.Recordset
Public adrsemp2 As New ADODB.Recordset
Public adrsumi As New ADODB.Recordset
'
Public adrsWOTGL As New ADODB.Recordset
''
Public adrsDept1 As New ADODB.Recordset
Public adrsPaid As New ADODB.Recordset
Public adrsRits As New ADODB.Recordset
Public adrsMod As New ADODB.Recordset
Public adrsDSR As New ADODB.Recordset
Public adrsOT As New ADODB.Recordset
Public adrsCO As New ADODB.Recordset
Public DBConn As New ADODB.Connection
Public adrsASC As New ADODB.Recordset


Public strAutoG As String

Public strVersionWithTital  As String
Public strVersionWithOutTital  As String
Public blnGoWithExportInExcell As Boolean

'original
'Public strTags As String
'new
Public strTags() As String

Public blnFlagForDept As Boolean
'Public strDatFilepath As String
Public Enum DataType
    NumericD
    StringD
End Enum
Public Type LeaveFile
    strLvInfo As String
    strLvTrn As String
    strLvBal As String
End Type

Public strSqlExport As String
Public Const strChecked = "þ"
Public Const strUnChecked = "q"
Public Const EmptyString As String = ""
Public Const strState = "KANPUR"
Public blnTagArray As Boolean
Public NewCapFlag  As Boolean   ' 06-08
Public LookFor As Variant
Public RepWith As Variant

Public SubLeaveFlag As Integer
Public PerAtt As Boolean
Public UserLocations As String
Public LVTRNFIELDS As Integer
Public Type SubMenu
    mnuCap As String
End Type
Public ExpSubMenu As SubMenu
Public FixedLvCode As String
