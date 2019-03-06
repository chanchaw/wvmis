VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CacheFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'窗体缓存
Private A_arrFrmCache() As dmFrm   '窗体缓存领域对象数组

Private A_LowIndex As Long   '最低INDEX
Private A_UpIndex As Long   '最高INDEX

Private strSQL As String

'根据对象编号获取开发环境中的对象即窗体
Private Function GetNewFrm(ByVal vObjectID As String) As dmFrm
    On Error GoTo IFERR
    Dim szLeft As String
    Dim rs As Recordset
    szLeft = Left$(vObjectID, 2)
    
    Select Case szLeft
        Case "11"
            strSQL = "SELECT B_ObjectID, B_FormName FROM G_BLS WHERE B_ObjectID='" & vObjectID & "'"
        Case "12"
            strSQL = "SELECT B_ObjectID, B_FormName FROM G_BL WHERE B_ObjectID='" & vObjectID & "'"
            Debug.Print strSQL
        Case "13"
            strSQL = "SELECT B_ObjectID, B_FormName FROM G_BLR WHERE B_ObjectID='" & vObjectID & "'"
        Case "19"
            strSQL = "SELECT B_ObjectID, B_FormName FROM G_CustomObject WHERE B_ObjectID='" & vObjectID & "'"
    End Select
    
    Set rs = New Recordset
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    Dim oDM As dmFrm
    Set oDM = New dmFrm
    With oDM
        .ObjectID = vObjectID
        Set .frm = GetFormNew(rs!B_FormName)
    End With
    
    Set GetNewFrm = oDM
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
IFERR:
    Dim szErr As String
    szErr = Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Function

'获取数组的最小和最大索引
Private Sub GetLUIndex()
    A_LowIndex = LBound(A_arrFrmCache)
    A_UpIndex = UBound(A_arrFrmCache)
End Sub

Private Sub AddFrm(ByVal vObjectID As String, ByRef oFrm As Object)
    '获取数组的上下限INDEX
    GetLUIndex
    
    ReDim Preserve A_arrFrmCache(A_LowIndex To A_UpIndex + 1)
    Dim oDM As dmFrm
    Set oDM = New dmFrm
    With oDM
        .ObjectID = vObjectID
        Set .frm = oFrm
    End With
    Set A_arrFrmCache(A_UpIndex + 1) = oDM
End Sub



Private Function JudgeExists(ByVal vObjectID As String) As Boolean
    JudgeExists = False
    
    '获取数组的上下限INDEX
    GetLUIndex
    
    Dim i As Long
    For i = A_LowIndex To A_UpIndex
        If A_arrFrmCache(i).ObjectID = vObjectID Then
            JudgeExists = True
            Exit Function
        End If
    Next
End Function

Private Function GetFrmfromArr(ByVal vObjectID As String) As Object
    '获取数组的上下限INDEX
    GetLUIndex
    
    Dim i As Long
    For i = A_LowIndex To A_UpIndex
        If A_arrFrmCache(i).ObjectID = vObjectID Then
            Set GetFrmfromArr = A_arrFrmCache(i).frm
            Exit Function
        End If
    Next
End Function

Private Sub PrintDM()
    GetLUIndex
    
    Dim i As Long
    Dim szTip As String
    For i = A_LowIndex To A_UpIndex
        
    Next
End Sub

'从数组中删除一个对象窗体
Public Sub DelFrm(ByVal vObjectID As String)
    
    GetLUIndex
    
    If A_UpIndex = 0 Then
        Exit Sub
    End If
    
    Dim i As Long, k As Long
    For i = A_LowIndex To A_UpIndex
        If A_arrFrmCache(i).ObjectID = vObjectID Then
            k = i
        End If
    Next
    
    For i = k + 1 To A_UpIndex
        Set A_arrFrmCache(i - 1) = A_arrFrmCache(i)
    Next
    
    ReDim Preserve A_arrFrmCache(A_LowIndex To A_UpIndex - 1)

End Sub

'获取指定对象编号的窗体
'1. 如果从数组中没有获取到对象窗体
'2. 那么创建新的窗体并且保存进数组中
Public Function GetFrm(ByVal vObjectID As String) As Object
    Dim oFrm As Object
    Set oFrm = GetFrmfromArr(vObjectID)
    
    '在缓存的数组中没有找到
    '则新建窗体
    If oFrm Is Nothing Then
        Set oFrm = GetNewFrm(vObjectID).frm
        AddFrm vObjectID, oFrm  '向数组中添加窗体
    End If
    
    Set GetFrm = oFrm
End Function

Private Sub Class_Initialize()
    ReDim A_arrFrmCache(0 To 0)
    
    Dim oDM As dmFrm
    Set oDM = New dmFrm
    With oDM
        .ObjectID = "000000"
    End With
    Set A_arrFrmCache(0) = oDM
End Sub
