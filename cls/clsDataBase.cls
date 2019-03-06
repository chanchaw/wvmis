VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'本类创建于2013年12月26日 22:42:30

Private strSQL As String
Private A_rsTable As New RecordSet

Private A_LogFile As String
Private A_Logger As New clsFile

'初始化本类的函数
'获取当前数据库下的所有用户表的名称
Public Sub initCls(ByVal vDBName As String)
    Set A_rsTable = New RecordSet
    strSQL = "Select * FROM " & vDBName & "..SysObjects Where XType='U' Order By Name"
    A_rsTable.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
End Sub

'判断一个数据表在当前的数据库中是否存在
'存在返回TRUE,否则返回FALSE
Public Function JudgeTableExist(ByVal vTableName As String) As Boolean
    Dim strTip As String
    
    JudgeTableExist = True
    
    If A_rsTable.State <> adStateOpen Then
        JudgeTableExist = False
        strTip = "数据库用户表记录集未能正常打开!" & vbNewLine
        strTip = strTip & "请先确认您是否对本类做了初始化!"
        MsgBox strTip, vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    If A_rsTable.RecordCount <= 0 Then
        JudgeTableExist = False
        strTip = "数据库用户表的数量为0!" & vbNewLine
        strTip = strTip & "请先确认您是否对本类做了初始化!"
        MsgBox strTip, vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    A_rsTable.Filter = " Name='" & vTableName & "'"
    If A_rsTable.RecordCount > 0 Then
        JudgeTableExist = True
    Else
        JudgeTableExist = False
    End If
    
    A_rsTable.Filter = ""
End Function

'判断一个表中是否存在某字段
Public Function JudgeFieldExist(ByVal vTableName As String, ByVal vFieldName As String) As Boolean
    Dim rs As New RecordSet
    
    '当表都不存在的时候,直接退出
    If JudgeTableExist(vTableName) = False Then
        JudgeFieldExist = False
        Exit Function
    End If
    
    
    Set rs = New RecordSet
    strSQL = "Select * FROM SysColumns Where id=Object_Id('" & vTableName & "')"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    rs.Filter = " Name='" & vFieldName & "'"
    If rs.RecordCount > 0 Then
        JudgeFieldExist = True
    Else
        JudgeFieldExist = False
    End If
    rs.Close
    Set rs = Nothing
    
End Function


'判断一个字符串中的所有字段是否存在于目标表中
'不存在的所有字段都给出提示
Public Function JudgeFieldsExist(ByVal vTableName As String, ByVal vFieldNames As String, ByVal vSymbol As String) As Boolean
    Dim rs As New RecordSet
    Dim szErrTip As String
    Dim curFieldName As String
    
    JudgeFieldsExist = True
    
    '当表都不存在的时候,直接退出
    If JudgeTableExist(vTableName) = False Then
        MsgBox "指定的表" & vTableName & "不存在！", vbOKOnly + vbInformation, "提示"
        JudgeFieldsExist = False
        Exit Function
    End If
    
    
    '当没有设置被检测的字段时候返回TRUE
    If Len(Trim$(vFieldNames)) <= 0 Then
        MsgBox "没有设置被检测的字段！", vbOKOnly + vbInformation, "提示"
        JudgeFieldsExist = False
        Exit Function
    End If
    
    
    '获取指定表下的所有字段
    Set rs = New RecordSet
    strSQL = "Select * FROM SysColumns Where id=Object_Id('" & vTableName & "')"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    If rs.RecordCount <= 0 Then
        MsgBox "指定的表下不存在任何字段！", vbOKOnly + vbInformation, "提示"
        rs.Close
        Set rs = Nothing
        JudgeFieldsExist = False
        Exit Function
    End If
        
    
    Dim szArr
    Dim i As Long
    szArr = Split(vFieldNames, vSymbol)
    
    
    
    '循环检测字段是否存在
    szErrTip = ""
    For i = 0 To UBound(szArr)
        curFieldName = Trim(szArr(i))
        rs.Filter = " Name='" & curFieldName & "'"
        
        If rs.RecordCount <= 0 Then
            szErrTip = szErrTip & curFieldName & vbNewLine
            JudgeFieldsExist = False
        End If
    Next

    '如果有不存在的字段那么给出提示
    If Len(Trim$(szErrTip)) > 0 Then
        szErrTip = "表" & vTableName & "中不存在字段：" & vbNewLine & szErrTip
        MsgBox szErrTip, vbOKOnly + vbInformation, "提示"
    End If
    
    rs.Close
    Set rs = Nothing
End Function

'vNoIdentity:默认不返回自增的列
Public Function GetTableFields(ByVal vTableName As String, _
    Optional ByVal vNoIdentity As Boolean = True) As String
    
    If vNoIdentity = True Then
        strSQL = "SELECT * FROM syscolumns AS s WHERE s.id=OBJECT_ID('" & vTableName & "') AND COLUMNPROPERTY(s.id,s.name,'IsIdentity')<>1"
    Else
        strSQL = "SELECT * FROM syscolumns AS s WHERE s.id=OBJECT_ID('" & vTableName & "')"
    End If
    Debug.Print strSQL
    Dim rs As New RecordSet
    Dim cls1 As New clsRecordset
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    GetTableFields = cls1.RecordSetToString(rs, "name", ",")
    
    rs.Close
    Set rs = Nothing
End Function

'执行外部SQL脚本文件
Public Function ExecSQLScript(ByVal vPath As String) As Boolean
    ExecSQLScript = True
    If Len(Trim$(vPath)) <= 0 Then
        ExecSQLScript = False
        MsgBox "指定的SQL脚本所在路径不存在！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    A_LogFile = App.Path & "\服务端升级日志.txt"
    A_Logger.WriteFileContent A_LogFile, "在" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "开始写服务端升级日志"
    A_Logger.WriteFileContent A_LogFile, "================================================================"



    If Right$(vPath, 1) = "\" Or Right$(vPath, 1) = "/" Then
        vPath = Left$(vPath, Len(vPath) - 1)
    End If
    
    '修改登录数据库的用户名默认数据库
    ChangeDBUserDefaultDBCur
     
    Dim rsFile As RecordSet
    Dim clsFile1 As clsFile
    
    Set clsFile1 = New clsFile
    Set rsFile = New RecordSet
    
    '获取SQL所在文件夹下的所有文件的绝对路径
    Set rsFile = clsFile1.TraversalFilesReturnRS(vPath)
    
    
    '给记录集排序，按照顺序执行SQL
    If rsFile.RecordCount <= 0 Then
        rsFile.Close
        Set rsFile = Nothing
        Exit Function
    End If
    
    
    rsFile.Sort = " B_Field0 ASC"
    rsFile.MoveFirst
    
    
    
    Dim rs As New RecordSet
    Do While Not rsFile.EOF
        '逐个文件执行SQL
        Set rs = New RecordSet
        strSQL = "master.dbo.xp_cmdshell ' osql -U " & Gm.SysID.DBInfo.DBUser & " -P " & Gm.SysID.DBInfo.DBPassWord & " -i """ & rsFile!B_Field0 & """'"
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        WriteLog rs
        
        rs.Close
        Set rs = Nothing
    
        rsFile.MoveNext
    Loop
    
    rsFile.Close
    Set rsFile = Nothing
    
End Function

'修改数据库用户默认的数据库
Public Sub ChangeDBUserDefaultDB(ByVal vUser As String, ByVal vDBName As String)
    strSQL = "sp_defaultdb '" & vUser & "','" & vDBName & "'"
    Gm.cnnTool.cnn.Execute strSQL
End Sub

'修改当前配置文件中指定的数据库登录用户
'的默认数据库为配置文件中指定的数据库
Public Sub ChangeDBUserDefaultDBCur()
    ChangeDBUserDefaultDB Gm.SysID.DBInfo.DBUser, Gm.SysID.DBInfo.DBName
End Sub



Private Sub WriteLog(ByRef vRs As RecordSet)
    If vRs.State <> adStateOpen Then
        Exit Sub
    End If
    
    If vRs.RecordCount <= 1 Then
        Exit Sub
    End If
    

    Dim szContent As String
    Do While Not vRs.EOF
        szContent = IIf(IsNull(vRs(0)), "", vRs(0))
        If Len(szContent) > 0 Then
            A_Logger.WriteFileContent A_LogFile, szContent
        End If
        vRs.MoveNext
    Loop
    
End Sub
