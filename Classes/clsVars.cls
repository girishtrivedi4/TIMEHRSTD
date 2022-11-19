VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public strShiftTmp As String            '' For the Shift
Public strHlsCode As String
Public strAbsCode As String
Public strPrsCode As String
Public strWosCode As String
Public strShiftOfDay As String           '' For Shift of the Day
Public strEmpCodeTrn As String           '' For Employee Code
Public strStatus As String              '' For Status
Public strDtTrn As String                '' For Processing Date as String
Public strIrrMark As String             '' For Irregular Mark
Public strRemarks As String             '' For Remarks per Record
Public sngPresent As Single             '' For The Present Display Based on Status
Public strAflg As String                '' For Aflg in Transaction
Public strDflg As String                '' For Dflg in Transaction
'change by  MIS2007DF020
Public bytTmpEnt As Integer                 '' For the Total Valid Entries Found
Public strLeaveStatus As String         '' For the Leave Status
Public strTmpLvtype As String           '' For the Leave type i.e Running or Working.
Public blnFoundPunches As Boolean       '' For Flag if Punches are Found or Not
Public blnTrnErl As Boolean             '' For Flag if Early Gone
Public blnTrnLate As Boolean            '' For Flag if Late Come
