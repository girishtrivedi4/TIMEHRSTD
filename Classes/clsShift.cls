VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public strShift As String   '' Shift Code
Public sngIN As Single      '' In Time
Public sngOut As Single     '' Out Time
Public sngHRS As Single     '' Shift Hours
Public blnNight As Boolean  '' Night Shift Flag
Public sngHalfE As Single   '' Half Day End
Public sngHalfS As Single   '' Second Half Start
Public sngB1I As Single     '' Break 1 in
Public sngB1O As Single     '' Break 1 Out
Public sngBH1 As Single     '' Break 1 Hours
Public sngB2I As Single     ''
Public sngB2O As Single     ''
Public sngBH2 As Single     ''
Public sngB3I As Single     ''
Public sngB3O As Single     ''
Public sngBH3 As Single     '' Till Break 3

Public sngUPTO As Single
''
