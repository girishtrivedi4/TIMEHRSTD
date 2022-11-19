VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public strEmp As String           '' Employee Code
Public strName As String          '' Emloyee Name
Public strCard As String          '' Employee Card
Public bytEntry As Byte           '' Entry Required
Public blnAuto As Boolean         '' Auto Shift Permission
Public strECat As String          '' Employee Category
Public strConv As String          '' Employee Conveyance Mode
Public strOff As String           '' Week Off
Public strOff2 As String          '' Week Off2
Public strWO13 As String          '' Week Off On 1_3
Public strWO24 As String          '' Week Off On 2_4
Public strShifttype As String     '' type of Shift
Public strEmpShift As String      '' Shift Code
Public dtJoin As Date             '' Employee Join Date
Public dtLeft As Date             '' Left date
'' For Details regarding Daily processing
Public bytWOHLAction As Byte      '' For Action on Week Off or Holiday
Public strAction3Shift As String  '' Assign this Shift on Week Off and Holiday
Public blnAutoOnPunch As Boolean  '' If Auto Shift if Punch Found
Public strActionBlank As String   '' If Blank Shift is found.

Public strAutoGroup As String
''
