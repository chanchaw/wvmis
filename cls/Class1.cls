VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevEnviroment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String

'根据对象编号判定对象属于：单表、单据、工程报表
'之后获取其字段设置的表名：G_BLSField、G_BLField、G_BLRField
Public Function GetObjectFieldTableName(ByVal vObjectID As String) As String
    Dim szPrefix As String
    szPrefix = Left$(vObjectID, 2)
    Select Case szPrefix
        Case "11"
            GetObjectFieldTableName = "G_BLSField"
        Case "12"
            GetObjectFieldTableName = "G_BLField"
        Case "13"
            GetObjectFieldTableName = "G_BLRField"
        Case Else
            GetObjectFieldTableName = ""
    End Select
End Function

'设置网格控件为TDBGrid的列宽
Public Sub SetColWidth8TDBGrid(ByVal vObjectID As String, ByRef vTDBGrid As TDBGrid)
    Dim i As Long
    Dim rs As New RecordSet
    Dim szTableName As String
    szTableName = GetObjectFieldTableName(vObjectID)
    
    strSQL = "SELECT * FROM " & szTableName & " WHERE B_ObjectID='" & vObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    For i = 0 To vTDBGrid.Columns.Count - 1
        rs.Filter = " B_FieldName='" & vTDBGrid.Columns(i).DataField & "'"
        If rs.RecordCount > 0 Then
            rs!B_GridWidth = vTDBGrid.Columns(i).width
            rs.Update
        End If
    Next
    
    rs.Close
    Set rs = Nothing
End Sub

'设置单表的列显示宽度
Public Sub SetBLSColWidth8TDBGrid(ByVal vObjectID As String, ByRef vTDBGrid As TDBGrid)
    Dim i As Long
    Dim rs As New RecordSet
    
    strSQL = "SELECT * FROM G_BLSField WHERE B_ObjectID='" & vObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    For i = 0 To vTDBGrid.Columns.Count - 1
        rs.Filter = " B_FieldName='" & vTDBGrid.Columns(i).DataField & "'"
        If rs.RecordCount > 0 Then
            rs!B_GridWidth = vTDBGrid.Columns(i).width
            rs.Update
        End If
    Next
    
    rs.Close
    Set rs = Nothing
End Sub

Public Function GetObjectTypeInt(ByVal vObjectID As String) As Long
    Dim szPrefix As String
    szPrefix = Left$(vObjectID, 2)
    
    Select Case szPrefix
        Case "11"
            GetObjectTypeInt = 3
        Case "12"
            GetObjectTypeInt = 4
        Case "13"
            GetObjectTypeInt = 5
    End Select
End Function

