VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dmSuitTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private A_DraftBill As String
Private A_DraftDetail As String
Private A_Bill As String
Private A_Detail As String

Private A_DraftMainKeyField As String
Private A_MainKeyField As String
Private A_DraftDetailKeyField As String
Private A_DetailKeyField As String

Private strSQL As String

Public Property Let DraftBill(ByVal vData As String)
    A_DraftBill = vData
    A_DraftMainKeyField = GetTableKeyField(A_DraftBill)
End Property

Public Property Get DraftBill() As String
    DraftBill = A_DraftBill
End Property

Public Property Let DraftDetail(ByVal vData As String)
    A_DraftDetail = vData
    A_DraftDetailKeyField = GetTableKeyField(A_DraftDetail)
End Property

Public Property Get DraftDetail() As String
    DraftDetail = A_DraftDetail
End Property

Public Property Let Bill(ByVal vData As String)
    A_Bill = vData
    A_MainKeyField = GetTableKeyField(A_Bill)
End Property

Public Property Get Bill() As String
    Bill = A_Bill
End Property

Public Property Let Detail(ByVal vData As String)
    A_Detail = vData
    A_DetailKeyField = GetTableKeyField(A_Detail)
End Property

Public Property Get Detail() As String
    Detail = A_Detail
End Property




Public Property Get DraftBillkey() As String
    DraftBillkey = A_DraftMainKeyField
End Property

Public Property Get Billkey() As String
    Billkey = A_MainKeyField
End Property

Public Property Get DraftDetailKey() As String
    DraftDetailKey = A_DraftDetailKeyField
End Property

Public Property Get DetailKey() As String
    DetailKey = A_DetailKeyField
End Property





'获取草稿主表的主键字段
Private Function GetTableKeyField(ByVal vTableName As String) As String
    Dim rs As New RecordSet
    strSQL = "Select top 1 * From " & vTableName
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Dim i As Long
    For i = 0 To rs.Fields.Count - 1
        If rs(i).Properties("KEYCOLUMN").Value = True Then
            GetTableKeyField = rs(i).name
        End If
    Next
    
    rs.Close
    Set rs = Nothing
End Function


