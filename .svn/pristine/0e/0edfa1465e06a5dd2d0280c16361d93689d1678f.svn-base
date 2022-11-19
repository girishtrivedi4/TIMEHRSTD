VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTataa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTata
' DateTime  : 10/04/2008 11:07
' Author    :
' Purpose   :
'---------------------------------------------------------------------------------------
Private mEmpcode As String
Private mEmpCat As String
Private mLast_Date As String
Private mDays As String
Private mEncash As String
Private mSD_01 As String
Private mSD_02 As String
Private mSD_03 As String
Private mTableName As String
'for empcode
Public Property Get Empcode() As String
    Empcode = mEmpcode
End Property

Public Property Let Empcode(ByVal vNewValue As String)
    mEmpcode = vNewValue
End Property

'for last date
Public Property Get Last_Date() As String
    Last_Date = mLast_Date
End Property

Public Property Let Last_Date(ByVal vNewValue As String)
    mLast_Date = vNewValue
End Property

'for days
Public Property Get Days() As String
    Days = mDays
End Property

Public Property Let Days(ByVal vNewValue As String)
    mDays = vNewValue
End Property

'for Encash
Public Property Get Encash() As String
    Encash = mEncash
End Property

Public Property Let Encash(ByVal vNewValue As String)
    mEncash = vNewValue
End Property

'for SD-01
Public Property Get SD_01() As String
    SD_01 = mSD_01
End Property

Public Property Let SD_01(ByVal vNewValue As String)
    mSD_01 = vNewValue
End Property

'for SD-02
Public Property Get SD_02() As String
    SD_02 = mSD_02
End Property

Public Property Let SD_02(ByVal vNewValue As String)
    mSD_02 = vNewValue
End Property

'for SD-03
Public Property Get SD_03() As String
    SD_03 = mSD_03
End Property

Public Property Let SD_03(ByVal vNewValue As String)
    mSD_03 = vNewValue
End Property

'for employee category
Public Property Get category() As String
    category = mEmpCat
End Property

Public Property Let category(ByVal vNewValue As String)
    mEmpCat = vNewValue
End Property

'for TableName
Public Property Get TableName() As String
    TableName = mTableName
End Property

Public Property Let TableName(ByVal vNewValue As String)
    mTableName = vNewValue
End Property

Public Sub Insert()
    AttendoCon.Execute "INSERT INTO " & TableName & _
    "(empcode,lst_date,days,encash,[SD-01],[SD-02],[SD-03]) VALUES ('" & _
    Empcode & "','" & Last_Date & "','" & FormateNumber(Days) & "','" & FormateNumber(Encash) & "','" & _
    FormateNumber(SD_01) & "','" & FormateNumber(SD_02) & "','" & FormateNumber(SD_03) & "')"
End Sub

Private Function FormateNumber(strInput As String) As String
    FormateNumber = Format(strInput, "00.00")
End Function
