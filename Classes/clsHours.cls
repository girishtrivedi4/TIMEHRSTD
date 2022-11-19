VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sngWorkHrs As Single     '' Total Work Hours
Public sngLateHrs As Single     '' Late Hours
Public sngEarlyHrs As Single    '' Early Hours
Public sngOverTime As Single    '' OT Hours
Public sngCOHrs As Single       '' Co Hours

Public sngRndLateHrs As Single     '' Round Off Late Hours '15-04
