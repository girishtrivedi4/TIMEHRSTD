VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeaveDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public strLvCode As String      '' Leave Code
Public strCat As String         '' Category
Public strLvName As String            '' Leave Description
Public publicsngQty As Single            '' Leave Quantity
Public sngAccQty As Single            '' Leave Accumulation Quantity
Public blnCarry As Boolean            '' Carry Balance Forward
Public blnCrImd As Boolean            '' Credit Immidiately or Next Year
Public blnFullPro As Boolean          '' Credit Full or Proportionately
Public blnLvType As Boolean           '' Leave Type
Public sngQty As Single
'this enum used to specify which leave file pass.
Public Enum eLeaveFile
    LvBal
    LvInfo
    LvTrn
End Enum

Public Enum eTrcd
    Opening = 1
    Credit
    Encash
    Avail
End Enum

