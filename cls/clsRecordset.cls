VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private strSQL As String


'将一个记录集中的内容生成为一个文本文档
Public Sub CreateRsFile(ByRef oRs As RecordSet)
    Dim rs As RecordSet
    Dim strLineData As String
    Dim i As Long
    Dim strFilter As String
    
    strFilter = "  |  "
    
    
    If oRs.State <> adStateOpen Then
        MsgBox "目标记录集未打开！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    If oRs.RecordCount <= 0 Then
        MsgBox "目标记录集的条目数为空！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    '1.克隆源记录集到本函数内的的临时记录集中
    Set rs = oRs.Clone
    
    
    
    '2.创建一个文本文档
    Dim fso As Object
    Dim strFileName As String
    strFileName = "记录集" & Format(Now, "YYYY-MM-DD HHMMSS") & ".txt"
    strFileName = App.Path & "\" & strFileName
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile strFileName
    
    
    '3.写入第一行字段名称
    strLineData = ""
    For i = 0 To rs.Fields.Count - 1
        strLineData = strLineData & rs(i).name & strFilter
    Next
    strLineData = Trim$(strLineData)
    strLineData = Left$(strLineData, Len(strLineData) - 1)
    WriteFileContent strFileName, strLineData
    
    
    '4.写入余下的数据
    rs.MoveFirst
    Do While Not rs.EOF
        strLineData = ""
        For i = 0 To rs.Fields.Count - 1
            strLineData = strLineData & rs(i).Value & strFilter
        Next
        WriteFileContent strFileName, strLineData
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    
    '5.打开上面创建的文本文档
    Dim a&
    a = ShellExecute(0, "open", strFileName, 0, 0, 1)
End Sub


'向文本文档追加写入文本（如果该文件不存在则创建并且写入）
'strFileName：文件名，绝对路径
'strWrite：要写入的内容
Private Sub WriteFileContent(ByVal strFileName As String, ByVal strWrite As String)
    Dim lngHandle As Long '句柄
    lngHandle = FreeFile() '取得句柄
    

    Open strFileName For Append As lngHandle    '打开文件
    Print #lngHandle, strWrite    '写入文本
    Close lngHandle    '关闭文件
End Sub

'将一个记录集中的数据写入到一个表中
Public Function WriteRsToDataBase(ByVal vTableName As String, ByRef vRs As RecordSet) As Boolean
    Dim rs1 As RecordSet
    
    If Len(Trim$(vTableName)) <= 0 Then
        WriteRsToDataBase = False
        Exit Function
    End If
    
    If vRs.State <> adStateOpen Then
        WriteRsToDataBase = False
        Exit Function
    End If
    
    If vRs.RecordCount <= 0 Then
        WriteRsToDataBase = False
        Exit Function
    End If
    
    Set rs1 = New RecordSet
    Set rs1 = vRs.Clone
    
    Dim rs2 As New RecordSet
    
    strSQL = "Select * From " & vTableName
    Set rs2 = New RecordSet
    rs2.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    
    rs1.MoveFirst
    Do While Not rs1.EOF
        rs2.AddNew
        
        rs1.movenext
    Loop
End Function

'本函数用于WriteRsToDataBase所以是非Public
Private Function GetTableRsCorrespond(ByVal vTableName As String) As RecordSet
    Dim rs As RecordSet
    Set rs = New RecordSet
    rs.Fields.Append "B_Field01", adVarChar
    rs.Fields.Append "B_Field02", adVarChar
    
    
End Function


Public Sub BuildRsFields(ByRef vRs As RecordSet)
    Dim rs As New RecordSet
    Dim rs1 As New RecordSet
    Dim rs2 As New RecordSet
    Dim lID01 As Long   '表G_BuildRsFieldBill的B_ID
    Dim lID02 As Long   '表G_BuildRsFieldDetail_Section的B_ItemID
    Dim lID03 As Long   '表G_BuildRsFieldDetail的B_ItemID
    
    
    strSQL = "Select * From G_BuildRsFieldBill Where B_GroupName='外发检验网格'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    lID01 = rs!B_id
    rs.Close
    Set rs = Nothing
    
    
    strSQL = "Select * From G_BuildRsFieldDetail_Section Where B_ID=" & lID01 & " Order By B_Order"
    Debug.Print strSQL
    Set rs1 = New RecordSet
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs1.RecordCount <= 0 Then
        rs1.Close
        Set rs1 = Nothing
        Exit Sub
    End If
    
    
    Set vRs = New RecordSet
    Do While Not rs1.EOF
        strSQL = "Select * From G_BuildRsFieldDetail Where B_ID=" & rs1!B_ItemID & " Order By B_Order"
        Debug.Print strSQL
        Set rs2 = New RecordSet
        rs2.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        
        Do While Not rs2.EOF
            '使用当前数据表G_BuildRsFieldDetail中罗列的字段进行填充记录集的字段
            If rs2!B_CurData = 1 Then
                vRs.Fields.Append rs2!B_FieldName, adVarChar, 100
            End If
            
            '使用外部数据表中的某字段上的一系列值进行填充
            If rs2!B_CurData = 0 Then
                Set rs = New RecordSet
                strSQL = "Select * From " & rs2!B_ExtendTable & " Where 1=1 Order By B_Order"
                Debug.Print strSQL
                rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
                
                Do While Not rs.EOF
                    vRs.Fields.Append rs(Trim$(rs2!B_ExtendField)), adVarChar, 100
                    rs.movenext
                Loop
                rs.Close
                Set rs = Nothing
            End If
            rs2.movenext
        Loop
        rs2.Close
        Set rs2 = Nothing
        
        rs1.movenext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    vRs.Open
End Sub


'将一个记录集转换为一个字符串
'当strFilter长度大于0时，每条记录间用该字符做间隔
'否则的话中间不做间隔
'strFieldName为目标字段的字段名
Public Function RecordSetToString(ByRef rs As RecordSet, ByVal strFieldName As String, ByVal strFilter As String) As String
    Dim str As String
       
    str = ""
    rs.MoveFirst
    Do While Not rs.EOF
        If Len(Trim(strFilter)) > 0 Then
            str = str & rs(strFieldName) & strFilter
        Else
            str = str & rs(strFieldName)
        End If
        rs.movenext
    Loop
       
    If Len(Trim(strFilter)) > 0 Then
        str = Left(str, Len(str) - 1)
    End If
       
    '形成字符串后，记录集自动移动到第一条记录上
    rs.MoveFirst
    RecordSetToString = str
End Function

'类似RecordSetToString的两层转换
Public Function RecordSetToStringMulti(ByRef vRs As RecordSet, _
    ByVal vDMMulti As domainMultiLevel) As String
        
    If vRs.State <> adStateOpen Then
        Exit Function
    End If
    
    If vRs.RecordCount <= 0 Then
        Exit Function
    End If
    
    Dim str As String
    Dim rs As New RecordSet
    Set rs = vRs.Clone
    
    str = ""
    rs.MoveFirst
    Do While Not rs.EOF
        '制作格式：第一层字段+第二层间隔符号+第二层字段+第一层间隔符号
        '即：vDMMulti.Field1+vDMMulti.Symb2+vDMMulti.Field2+vDMMulti.Symb1
        str = str & rs(vDMMulti.Field1) & vDMMulti.Symb2
        str = str & rs(vDMMulti.Field2) & vDMMulti.Symb1
        rs.movenext
    Loop
       
    str = Left(str, Len(str) - 1)

    RecordSetToStringMulti = str
End Function


'逐行循环vRs2的记录
'在vRs1中将字段vField1上VALUE和vRs2一样的替换为某一字符串（这里采用循环可清理掉N多条相同的记录）
'使用vReplaceStr来替换
Public Function replaceRecord(ByRef vRs1 As RecordSet, ByRef vRs2 As RecordSet, _
    ByVal vField1 As String, ByVal vFiedl2 As String, ByVal vReplaceStr As String)
    
    Dim szTemp As String

    If vRs1.State <> adStateOpen Then
        Exit Function
    End If
    
    If vRs2.State <> adStateOpen Then
        Exit Function
    End If
    
    If vRs1.RecordCount <= 0 Then
        Exit Function
    End If
    
    If vRs2.RecordCount <= 0 Then
        Exit Function
    End If
    
    
    vRs2.MoveFirst
    Do While Not vRs2.EOF
        vRs1.Filter = " " & vField1 & "='" & vRs2(Trim$(vFiedl2)) & "'"
        If vRs1.RecordCount > 0 Then
            Do While Not vRs1.EOF
                vRs1(Trim$(vField1)) = vReplaceStr
                vRs1.movenext
            Loop
        End If
        vRs2.movenext
    Loop
    
    
    vRs1.Filter = ""
    If vRs1.RecordCount > 0 Then
        vRs1.MoveFirst
        Do While Not vRs1.EOF
            szTemp = IIf(IsNull(vRs1(Trim$(vField1))), "", vRs1(Trim$(vField1)))
            If Len(Trim$(szTemp)) <= 0 Then
                vRs1.delete
                vRs1.Update
            End If
            vRs1.movenext
        Loop
    End If
End Function

Public Sub FillUnConnectRecordSet(ByRef sRs As RecordSet, ByRef tRs As RecordSet, _
    Optional ByVal vFillData As Boolean = True)
    
    On Error Resume Next
    Dim i As Long
       
    Set tRs = New RecordSet
    For i = 0 To sRs.Fields.Count - 1
        tRs.Fields.Append sRs.Fields(i).name, sRs.Fields(i).Type, sRs.Fields(i).DefinedSize, sRs.Fields(i).Attributes
    Next
       
    tRs.Open
    
    '填充数据
    If vFillData = True Then
        Do While Not sRs.EOF
            tRs.AddNew
            For i = 0 To sRs.Fields.Count - 1
                tRs.Fields(Trim(sRs(i).name)).Value = IIf(IsNull(sRs.Fields(Trim(sRs.Fields(i).name)).Value), Null, sRs.Fields(Trim(sRs.Fields(i).name)).Value)
            Next
            tRs.Update
            sRs.movenext
        Loop
          
        tRs.MoveFirst
    
    End If
End Sub

'判断两个记录集在某字段上是否完全一致（包括记录数和指定字段下的VALUE）
'完全一样的话则返回TRUE，否则返回FALSE
Public Function JudgeTheSame(ByVal vRs1 As RecordSet, _
    ByVal vRs2 As RecordSet, ByVal vField1 As String, _
    ByVal vField2 As String) As Boolean

    If vRs1.State <> adStateOpen Then
        JudgeTheSame = False
        Exit Function
    End If
    
    If vRs2.State <> adStateOpen Then
        JudgeTheSame = False
        Exit Function
    End If
    
    If vRs1.RecordCount <= 0 Then
        JudgeTheSame = False
        Exit Function
    End If
    
    If vRs2.RecordCount <= 0 Then
        JudgeTheSame = False
        Exit Function
    End If
    
    vRs1.Filter = ""
    vRs2.Filter = ""
    
    If vRs1.RecordCount <> vRs2.RecordCount Then
        JudgeTheSame = False
        Exit Function
    End If
    
    vRs1.Sort = " " & vField1 & " asc"
    vRs2.Sort = " " & vField2 & " asc"
    
    vRs1.MoveFirst
    vRs2.MoveFirst
    
    Do While Not vRs1.EOF
        If vRs1(Trim$(vField1)) <> vRs2(Trim$(vField2)) Then
            JudgeTheSame = False
            Exit Function
        End If
        
        vRs1.movenext
        vRs2.movenext
    Loop
    
    JudgeTheSame = True
End Function

'返回一个表的所有字段，自增的字段除外
Public Function getTableFieldsWithoutIncream(ByVal vTableName As String) As String
    Dim strSQL As String
    Dim rs As RecordSet
    Dim i As Long
    Dim szReturn As String
    
    Set rs = New RecordSet
    strSQL = "Select top 1 * from " & vTableName

    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.Fields.Count <= 0 Then
        getTableFieldsWithoutIncream = ""
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).name
        If rs.Fields(i).Properties("ISAUTOINCREMENT").Value = False Then
            szReturn = szReturn & rs.Fields(i).name & ","
        End If
    Next
    
    szReturn = Left$(szReturn, Len(szReturn) - 1)
    rs.Close
    Set rs = Nothing
    
    getTableFieldsWithoutIncream = szReturn
End Function

'获取一个字段在其所在的记录集中的列序号，如果返回-1表示不存在
Public Function GetColIndex(ByRef vRs As RecordSet, ByVal vFieldName As String) As Long
    Dim rs As New RecordSet
    Set rs = vRs.Clone
    Dim i As Long
    
    
    GetColIndex = -1
    For i = 0 To rs.Fields.Count
        If rs(i).name = vFieldName Then
            GetColIndex = i
            Exit Function
        End If
    Next
End Function

Public Sub PrintFields(ByRef vRs As RecordSet)
    Dim rs As New RecordSet
    Dim i As Long
    Dim szContent As String
    
    Set rs = vRs.Clone
    For i = 0 To rs.Fields.Count - 1
        szContent = ""
        szContent = rs(i).name & "," & rs(i).Type
        Debug.Print szContent
    Next
    
    rs.Close
    Set rs = Nothing
End Sub

'将一个字符串通过固定的间隔符号转换为一个记录集返回
'vString：有固定间隔符号的一个字符串
'vFilterSymbol：间隔符号
'vFieldName：形成的记录集中字段的名称 - 生成的记录集只有一个字段
Public Function String2RS(ByVal vString As String, ByVal vFilterSymbol As String, _
    ByVal vFieldName As String) As RecordSet
    
    Dim i As Long
    Dim rs As New RecordSet
    
    vFieldName = Trim$(vFieldName)
    rs.Fields.Append vFieldName, adVarChar, 100
    rs.Open
    
    Dim arrString() As String
    arrString = Split(vString, vFilterSymbol)
    For i = 0 To UBound(arrString)
        rs.AddNew
        rs(vFieldName) = arrString(i)
    Next
    
    Set String2RS = rs.Clone
    rs.Close
    Set rs = Nothing
End Function

'在二维表格上循环打印数据
Public Sub PrintData(ByRef vRs As RecordSet)
    Dim rs As New RecordSet
    Dim i As Long
    Dim szTemp As String
    
    Set rs = vRs.Clone
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    '1.打印字段
    szTemp = ""
    For i = 0 To rs.Fields.Count - 1
        szTemp = szTemp & rs.Fields(i).name & ","
    Next
    szTemp = Left$(szTemp, Len(szTemp) - 1)
    
    Debug.Print szTemp
    
    
    '2.打印数据
    rs.MoveFirst
    Do While Not rs.EOF
        szTemp = ""
        For i = 0 To rs.Fields.Count - 1
            szTemp = szTemp & rs.Fields(i).Value & ","
        Next
        szTemp = Left$(szTemp, Len(szTemp) - 1)
        Debug.Print szTemp
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub


Public Function GetRS(ByVal vSQL As String, Optional ByVal vCanUpdate As Boolean = True) As RecordSet
    Dim rs As New RecordSet
    If vCanUpdate = True Then
        rs.Open vSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Else
        rs.Open vSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    End If
    
    Set GetRS = rs.Clone
    rs.Close
    Set rs = Nothing
End Function

'在记录集中判断一个字段是否存在
Public Function JudgeFieldExists(ByRef vRs As RecordSet, ByVal vFieldName As String) As Boolean
    JudgeFieldExists = False
    
    If vRs.State <> adStateOpen Then
        JudgeFieldExists = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To vRs.Fields.Count - 1
        If vRs.Fields(i).name = vFieldName Then
            JudgeFieldExists = True
        End If
    Next
End Function


Public Function GetFieldValue(ByRef vField As ADODB.Field) As String
    Select Case vField.Type
        Case adVarChar, adChar
            GetFieldValue = IIf(IsNull(vField.Value), "", vField.Value)       ' 1--String型
        Case adInteger, adSingle, adSmallInt, adDouble, adNumeric
            GetFieldValue = IIf(IsNull(vField.Value), 0, vField.Value)   ' 2--数值型
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            GetFieldValue = IIf(IsNull(vField.Value), Format(Now, "YYYY-MM-DD"), Format(vField.Value, "YYYY-MM-DD"))   ' 3--日期型
        Case Else
            GetFieldValue = IIf(IsNull(vField.Value), "", vField.Value)       ' 1--String型
    End Select
End Function

'判断一个记录集中的某字段上的值是否唯一
'唯一则返回TRUE，否则返回FALSE
Public Function JudgeFieldValueUnique(ByRef vRs As RecordSet, _
    ByVal vField As String) As Boolean
    
    
    JudgeFieldValueUnique = True
    If vRs.State <> adStateOpen Then
        JudgeFieldValueUnique = True
        Exit Function
    End If
    
    If vRs.RecordCount <= 0 Then
        JudgeFieldValueUnique = True
        Exit Function
    End If
    
    Dim rs As New RecordSet
    Dim szTemp As String
    Dim szChecked As String
    
    Set rs = vRs.Clone

    Do While Not rs.EOF
        If Len(szTemp) <= 0 Then
            szTemp = rs(Trim$(vField))
        Else
            szChecked = rs(Trim$(vField))
            If szChecked <> szTemp Then
                JudgeFieldValueUnique = False
                Exit Function
            End If
        End If
        
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Function

'返回构建了字段的没有数据的记录集
Public Function buildRsNoData(ByRef vRs As RecordSet) As RecordSet
'    Dim rs As New RecordSet
'    Dim i As Long
'    For i = 0 To vRs.Fields.Count - 1
'        rs.Fields.Append vRs.Fields(0).name, vRs.Fields(0).Type, vRs.Fields(0).DefinedSize, vRs.Fields(0).Attributes
'    Next
'    rs.Open
'
'    buildRsNoData = rs.Clone
End Function


'填充有筛选条件的记录集
Public Function buildRsWithData(ByRef vRs As RecordSet, ByVal vFilterString As String) As RecordSet
    Dim rs As New RecordSet
    Dim i As Long
    For i = 0 To vRs.Fields.Count - 1
        rs.Fields.Append vRs.Fields(i).name, vRs.Fields(i).Type, vRs.Fields(i).DefinedSize, vRs.Fields(i).Attributes
    Next
    rs.Open
    
    vRs.Filter = vFilterString
    Do While Not vRs.EOF
        rs.AddNew
        For i = 0 To vRs.Fields.Count - 1
            rs(i) = vRs(i)
        Next
        vRs.movenext
    Loop
    
   Set buildRsWithData = rs.Clone
End Function

