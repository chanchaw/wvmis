VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'下面的Excel必须引用：
'Microsoft Excel 11.0 Object Library
'一般路径为：X\Program File\Microsoft Office\OFFICE11\EXCEL.EXE

'Dim xlApp As Excel.Application '定义EXCEL类
'Dim xlBook As Excel.Workbook '定义工件簿类
'Dim xlSheet As Excel.Worksheet '定义工作表类

Dim xlApp As Object '定义EXCEL类
Dim xlBook As Object '定义工件簿类
Dim xlSheet As Object '定义工作表类

Private strSQL As String
Private A_rsRules As New RecordSet
Private A_rsMappingDetail As New RecordSet
Public W_LogFileName As String
Private A_TableName As String   '被导入数据的表名

Private clsLog As New clsFile


'本类为文件操作相关类


'创建批处理文件
'第一个参数为绝对路径的文件名
'第二个参数为批处理文件的内容
Public Sub CreateFileBat(ByVal m_FileName As String, ByVal m_Content As String)
    
    Open m_FileName For Append As #1
    Print #1, m_Content
    Close #1
    
End Sub


'向文本文档追加写入文本（如果该文件不存在则创建并且写入）
'strFileName：文件名，绝对路径
'strWrite：要写入的内容
Public Sub WriteFileContent(ByVal strFileName As String, ByVal strWrite As String)
    Dim lngHandle As Long '句柄
    lngHandle = FreeFile() '取得句柄
    

    Open strFileName For Append As lngHandle    '打开文件
    Print #lngHandle, strWrite    '写入文本
    Close lngHandle    '关闭文件
End Sub

'向文本文档追加写入文本（如果该文件不存在则创建并且写入）
'strFileName：文件名，绝对路径
'strWrite：要写入的内容
'vWriteTo：1表示要写入，0表示不要写入
Public Sub WriteFileContent_WithSwitch(ByVal strFileName As String, _
    ByVal strWrite As String, ByVal vWriteTo As Long)
    
    If vWriteTo = 0 Then
        Exit Sub
    End If
    
    Dim lngHandle As Long '句柄
    lngHandle = FreeFile() '取得句柄
    

    Open strFileName For Append As lngHandle    '打开文件
    Print #lngHandle, strWrite    '写入文本
    Close lngHandle    '关闭文件
End Sub


'打开一个文件
'vFileName：文件名（绝对路径）
Public Sub OpenFile(ByVal vFileName As String)
    Dim a&
    a = ShellExecute(0, "open", vFileName, 0, 0, 1)
End Sub


'判断某文件是否存在
'不存在返回false
'存在的话返回true
Public Function JudgeFileExist(ByVal m_FileName As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    JudgeFileExist = fso.FileExists(m_FileName)
       
    
    Set fso = Nothing
End Function


'删除文件
Public Sub DeleteFile(ByVal m_FileName As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFile (m_FileName)
End Sub


'清空一个文本文档中的内容
Public Sub ClearTxtContent(ByVal vFileName As String)
    Open vFileName For Output As #1
    Close #1
End Sub


'逐行从文本文档中读取内容
Public Function ReadTXT(TXTPath As String) As String
    Dim TXTCon, str As String, i As Long
    
    Dim lngHandle As Long '句柄
    lngHandle = FreeFile() '取得句柄
    
    
    Open TXTPath For Input As lngHandle
    Do While Not EOF(lngHandle)
        'Line Input lngHandle,TXTCon
        Line Input #lngHandle, TXTCon
        'Str = Str & TXTCon & vbCrLf
        str = str & TXTCon & vbNewLine
    Loop
    Debug.Print str
    ReadTXT = str
    
    Close lngHandle
End Function

'逐行读取文本文档中的内容，每行之间使用英文的逗号间隔
Public Function ReadTXT_InsertSym(TXTPath As String) As String
    Dim TXTCon, str As String, i As Long
    
    Dim lngHandle As Long '句柄
    lngHandle = FreeFile() '取得句柄
    
    
    Open TXTPath For Input As lngHandle
    Do While Not EOF(lngHandle)
        Line Input #lngHandle, TXTCon
        str = str & TXTCon & ","
    Loop
    ReadTXT_InsertSym = Left(str, Len(str) - 1)
    
    Close lngHandle
End Function


'从文本文档中读取最后一行的内容
'参数为文件的全路径文件名
Private Function ReadTXTLast(ByVal vFileName As String) As String
    Dim TXTCon, str As String, i As Long

    Open vFileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, TXTCon
    Loop
    ReadTXTLast = TXTCon
End Function

'通过弹出打开文件对话框来获取一个文件的绝对路径
'返回文件的绝对路径
Public Function OpenFile_Dialog() As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = "Excel文件 (*.xls)|*.xls|所有文件(*.*)|*.*|Excel文件 (*.xlsx)|*.xlsx"
        .ShowOpen
        szFileName = .FileName
    End With
    
    OpenFile_Dialog = Trim(szFileName)
    
    Set o = Nothing
End Function


Public Function OpenFile_DialogAllEt() As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = "所有文件(*.*)|*.*"
        .ShowOpen
        szFileName = .FileName
    End With
    
    OpenFile_DialogAllEt = Trim(szFileName)
    
    Set o = Nothing
End Function


'读取一个EXCEL文件中的内容到记录集中
'传递进来的参数是文件的绝对路径（全文件名）
'vRows:在EXCEL文件中读取的行数
'vCols:在EXCEL文件中读取的列数
Public Function ReadExcelToRecordSet(ByVal vFileName As String, _
    ByVal vRows As Long, ByVal vCols As Long, ByVal vExcludeRows As Long, _
    ByVal vExcludeCols As Long) As RecordSet
    
    On Error GoTo IFERR
    
    Dim i As Long, J As Long
    Dim retRecordSet As RecordSet
    Dim szErrExcelFileName As String
    Dim szTemp As String
    Dim szUnit As String
    Dim clsString1 As New clsString
    
    If Len(Trim$(vFileName)) <= 0 Then
        Exit Function
    End If
    
    If vRows <= 0 Then
        Exit Function
    End If
    
    If vCols <= 0 Then
        Exit Function
    End If
    
    '获取错误文件的绝对路径+文件名称
    '在2015年10月8日 13:25:01打卷将错误的数据导入到一个EXCEL文件中记录下来
    '后来由于添加了on error resume next，不可实现。遂将下面的一个代码块注释掉。
'    szErrExcelFileName = Format(Now, "YYYYMMDD HHMMSS") & "导入失败的数据.xls"
'    If Right$(g_ErrFilePath, 1) <> "\" Then
'        g_ErrFilePath = g_ErrFilePath & "\"
'    End If
'    szErrExcelFileName = g_ErrFilePath & szErrExcelFileName
    
    
    writeLogImportLog "开始创建Excel应用程序"
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL应用类
    xlApp.Visible = False '设置EXCEL不可见
    
    Set xlBook = xlApp.Workbooks.Open(vFileName) '打开EXCEL工作簿
    writeLogImportLog "工作簿打开完毕"
    Set xlSheet = xlBook.Worksheets(1) '打开EXCEL工作表
    writeLogImportLog "sheet打开完毕"
    
    Set retRecordSet = New RecordSet
    For i = 1 To vCols
        retRecordSet.Fields.Append "B_Field" & i, adVarChar, 100
    Next
    retRecordSet.Open
    writeLogImportLog "空记录集初始化完毕"
    
    writeLogImportLog "导入的数据内容如下：////////////////"
    For i = (1 + vExcludeRows) To (vRows + vExcludeRows)
        retRecordSet.AddNew
        szTemp = ""
        For J = (1 + vExcludeCols) To (vCols + vExcludeCols)
            szUnit = clsString1.KillCrlfSpace(xlSheet.Cells(i, J))
            retRecordSet(J - 1) = szUnit
            szTemp = szTemp & szUnit & ","
        Next
        szTemp = Left$(szTemp, Len(szTemp) - 1)
        writeLogImportLog szTemp
    Next
    writeLogImportLog "数据导入完毕，等待关闭/////////////"
    
    
    xlBook.Close (True) '关闭EXCEL工作簿
    xlApp.Quit '关闭EXCEL
    Set xlApp = Nothing '释放EXCEL对象
    
    writeLogImportLog "隐式关闭完毕，等待返回"
    Set ReadExcelToRecordSet = retRecordSet.Clone
    
    Exit Function
IFERR:
    Dim szErr As String
    szErr = "在ReadExcelToRecordSet中出错！" & vbNewLine & Err.Description
    writeLogImportLog szErr
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Function

'根据表G_ExcelMappingTableMain和G_ExcelMappingTableDetail
'将传入的记录集中的数据写入到对应的表中
'oRs:传入的记录集，其中带有数据
'vMappingIndex：对应于数据表G_ExcelMappingTableMain中B_ID。即按照这个编号的设定来导入到某个表中
'并且按照G_ExcelMappingTableDetail中设定的字段对应规则
'要使用本函数参照“田村精工包装打印系统”中的3个表的套表：
'G_ExcelMappingTableMain，G_ExcelMappingTableDetail，G_ExcelMappingTableRules
Private Function WriteRSToTable_ByMappingTable(ByRef oRs As RecordSet, ByVal vMappingIndex As Long) As Boolean
    On Error Resume Next
    Dim rsMap As RecordSet
    Dim rsData As RecordSet
    Dim rsTable As RecordSet
    Dim i As Long
    
    WriteRSToTable_ByMappingTable = True
    

    
    '获取表G_ExcelMappingTableMain，G_ExcelMappingTableDetail指定B_ID的设置
    strSQL = "exec dbo.P_GetMappingTable " & vMappingIndex
    Debug.Print strSQL
    Set rsMap = New RecordSet
    rsMap.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rsMap.RecordCount <= 0 Then
        WriteRSToTable_ByMappingTable = False
        rsMap.Close
        Set rsMap = Nothing
        Exit Function
    End If
    
    '获取即将被导入数据的表名称
    A_TableName = IIf(IsNull(rsMap!B_TableName), "", rsMap!B_TableName)
    
    If oRs.State <> adStateOpen Then
        WriteRSToTable_ByMappingTable = False
        rsMap.Close
        Set rsMap = Nothing
        Exit Function
    End If
    
    If oRs.RecordCount <= 0 Then
        WriteRSToTable_ByMappingTable = False
        rsMap.Close
        Set rsMap = Nothing
        Exit Function
    End If
    
    Set rsData = New RecordSet
    Set rsData = oRs.Clone
    
    
    Set rsTable = New RecordSet
    strSQL = "Select * From " & rsMap!B_TableName
    Debug.Print strSQL; strSQL
    rsTable.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    'rsData:从EXCEL读取的数据
    'rsTable：数据导入的目的表中获取的记录集
    rsData.MoveFirst
    Do While Not rsData.EOF
        '首先判断某字段上的值唯一的情况。
        '如果需要判断，同时判断得出没有重复记录则继续ADDNEW
        '如果需要判断，同时判断得出有重复的记录，则跳到下一个EXCEL中的记录继续执行。
        If JudgeRules_Repeat(vMappingIndex, rsData, rsTable) = True Then
    
            rsTable.AddNew
            For i = 1 To rsData.Fields.Count
                rsMap.Filter = " B_ExcelColIndex='" & i & "'"
                If rsMap.RecordCount > 0 Then
                    rsTable(Trim(rsMap!B_FieldName)) = rsData(i - 1)
                End If
            Next
            
            '如果当前行不符合条件规则，则要取消掉
            If JudgeRules(vMappingIndex, rsTable) = False Then
                rsTable.CancelUpdate
            Else
                rsTable.Update
            End If
        
        End If
        
        rsData.movenext
    Loop
    
    rsTable.Close
    Set rsTable = Nothing
    
    
    rsData.Close
    Set rsData = Nothing
    
    rsMap.Close
    Set rsMap = Nothing
End Function

'通过弹出一个窗体，在其上面设置源Excel中的导入数据的行数，列数，
'排除的前N行，排除的前N列。来返回一个导入数据的区域
'返回的结构体在modStruecture中声明
'这里使用到的结构体位于soft/公用类/modStructure
Private Function getImportExcelRect() As importExcelRect
    Dim frm1 As New frmSetGridRowsCols
    Dim rtn As importExcelRect
    frm1.Show vbModal
    
    rtn.Rows = frm1.w_Rows
    rtn.cols = frm1.w_Cols
    rtn.excludeRows = frm1.w_ExcludeRows
    rtn.excludeCols = frm1.w_ExcludeCols
    
    getImportExcelRect = rtn
    
    Unload frm1
End Function

Public Function WriteRSToTable() As Boolean
    Dim szFileName As String
    Dim rsExcel As RecordSet
    
    Dim i As Long, J As Long, m As Long, n As Long
    
    Dim szTemp As String
    
    WriteRSToTable = True
    
    szFileName = OpenFile_Dialog
    If Len(Trim(szFileName)) <= 0 Then
        WriteRSToTable = False
        MsgBox "文件名称非法！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    

    '获取Excel文件中数据源区域
    writeLogImportLog "在" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "开始的导入操作======================="
    writeLogImportLog "开始打开数据源区域rect"
    
    Dim RECT As importExcelRect
    RECT = getImportExcelRect
    i = RECT.Rows
    J = RECT.cols
    m = RECT.excludeRows
    n = RECT.excludeCols
    
    writeLogImportLog "获取的数据源区域为：行数=" & i & ",列数=" & J & ",排除前" & m & "行,排除前" & n & "列"
    
    
    
    If i <= 0 Or J <= 0 Then
        WriteRSToTable = False
        MsgBox "行数或者列数非法！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    writeLogImportLog "开始进入ReadExcelToRecordSet"
    Set rsExcel = ReadExcelToRecordSet(szFileName, i, J, m, n)
    writeLogImportLog "入ReadExcelToRecordSet执行完毕！"
    
    WriteRSToTable_ByMappingTable rsExcel, 1
    writeLogImportLog "WriteRSToTable_ByMappingTable执行完毕"
End Function


'昌盛导入考勤数据用到.导入之前先删除之前的所有数据
Public Function WriteRSToTable_ByGroupName(ByVal vGroupName As String) As Boolean
    Dim szFileName As String
    Dim rsExcel As RecordSet
    
    Dim i As Long, J As Long, m As Long, n As Long
    
    WriteRSToTable_ByGroupName = True
    
    szFileName = OpenFile_Dialog
    If Len(Trim(szFileName)) <= 0 Then
        WriteRSToTable_ByGroupName = False
        MsgBox "文件名称非法！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    
    Dim frm1 As New frmSetGridRowsCols
    'frm1.W_ExcelImportDefaultGroupName = vGroupName   '田村精工没有设置表G_ExcelImportDefault.昌盛考勤系统有
    frm1.Show vbModal
    
    i = frm1.w_Rows
    J = frm1.w_Cols
    m = frm1.w_ExcludeRows
    n = frm1.w_ExcludeCols
    
    Unload frm1
    
    If i <= 0 Or J <= 0 Then
        WriteRSToTable_ByGroupName = False
        MsgBox "行数或者列数非法！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    '先删除之前的所有数据
    DelBeforeImport vGroupName
    
    Set rsExcel = ReadExcelToRecordSet(szFileName, i, J, m, n)
    
    WriteRSToTable_ByMappingTable rsExcel, 1
End Function

'在导入数据前把明细表中的数据全部删除
Private Sub DelBeforeImport(ByVal vGroupName As String)
    Dim rs As RecordSet
    strSQL = "Select * From G_ExcelImportDefault Where B_GroupName='" & vGroupName & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If IIf(IsNull(rs!B_DelBeforeImport), 0, rs!B_DelBeforeImport) = 1 Then
        strSQL = "Truncate table G_KQDetail"
        Gm.cnnTool.cnn.Execute strSQL
    End If
    rs.Close
    Set rs = Nothing
End Sub


'返回TRUE表示，当前一行数据可以进行保存
'返回FALSE表示当前一行数据需要通过CANCLE取消掉
Private Function JudgeRules(ByVal vMappingIndex As Long, ByRef rsTable As RecordSet) As Boolean
    On Error GoTo IFERR
    Dim rsRules As RecordSet
    Dim szGZFieldName As String
    Dim szGZFieldValue As String
    Dim szFilterString As String
    
    JudgeRules = True
    
    Set rsRules = New RecordSet
    strSQL = "Select * From G_ExcelMappingTableRules Where B_ID=" & vMappingIndex
    rsRules.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rsRules.RecordCount <= 0 Then
        rsRules.Close
        Set rsRules = Nothing
        Exit Function
    End If
    
    Do While Not rsRules.EOF
        '数值大于0
        If IIf(IsNull(rsRules!B_SZDY0), 0, rsRules!B_SZDY0) = 1 Then
            If IIf(IsNull(rsTable(Trim$(rsRules!B_FieldName))), 0, rsTable(Trim$(rsRules!B_FieldName))) <= 0 Then
                JudgeRules = False
                rsRules.Close
                Set rsRules = Nothing
                Exit Function
            End If
        End If
        
        '长度大于0
        If IIf(IsNull(rsRules!B_CDDY0), 0, rsRules!B_CDDY0) = 1 Then
            If Len(Trim(IIf(IsNull(rsTable(Trim$(rsRules!B_FieldName))), "", rsTable(Trim$(rsRules!B_FieldName))))) <= 0 Then
                JudgeRules = False
                rsRules.Close
                Set rsRules = Nothing
                Exit Function
            End If
        End If
        
        
        '某字段上的值唯一
'        If IIf(IsNull(rsRules!B_OnlyOne), 0, rsRules!B_OnlyOne) = 1 Then
'            szGZFieldName = Trim$(rsRules!B_FieldName)
'            szGZFieldValue = rsTable(Trim$(rsRules!B_FieldName))
'            szFilterString = " " & szGZFieldName & "='" & szGZFieldValue & "'"
'            Debug.Print szFilterString
'            rsTable.Filter = szFilterString
'            If rsTable.RecordCount > 0 Then
'                JudgeRules = False
'                rsRules.Close
'                Set rsRules = Nothing
'                Exit Function
'            End If
'        End If
        
        
        rsRules.movenext
    Loop
    
    rsRules.Close
    Set rsRules = Nothing
    
    Exit Function
IFERR:
    Dim szErr As String
    szErr = Err.Description & vbNewLine
    szErr = szErr & "发生于JudgeRules！"
    MsgBox Err.Description, vbOKOnly + vbInformation, "提示"
End Function

'从EXCEL表导入产品情况数据
Public Function ImportExcelData_Goods() As Boolean
    Dim szFileName As String
    Dim rsExcel As RecordSet
    
    Dim i As Long, J As Long, m As Long, n As Long
    
    ImportExcelData_Goods = True
    
    szFileName = OpenFile_Dialog
    If Len(Trim(szFileName)) <= 0 Then
        ImportExcelData_Goods = False
        MsgBox "文件名称非法！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    
    Dim frm1 As New frmSetGridRowsCols
    frm1.Show vbModal
    
    i = frm1.w_Rows
    J = frm1.w_Cols
    m = frm1.w_ExcludeRows
    n = frm1.w_ExcludeCols
    
    Unload frm1
    
    If i <= 0 Or J <= 0 Then
        ImportExcelData_Goods = False
        MsgBox "行数或者列数非法！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    
    Set rsExcel = ReadExcelToRecordSet(szFileName, i, J, m, n)
    
    WriteRSToTable_ByMappingTable rsExcel, 2
End Function

'判断规则表G_ExcelMappingTableRules中列B_OnlyOne
'oRsMap:传递进来rsMap，即Excel和数据表之间的映射表
Private Function JudgeOnlyOne(ByRef oRsMap As RecordSet, ByRef oRsData As RecordSet) As Boolean
    Dim rsMap As RecordSet
    Set rsMap = oRsMap.Clone
    
    Dim rsData As RecordSet
    Set rsData = oRsData.Clone
    
    If rsMap.State <> adStateOpen Then
        JudgeOnlyOne = True
        Exit Function
    End If
    
    If rsMap.RecordCount <= 0 Then
        JudgeOnlyOne = True
        Exit Function
    End If
    
    rsMap.MoveFirst
    Do While Not rsMap.EOF
        If IIf(IsNull(rsMap!B_OnlyOne), 0, rsMap!B_OnlyOne) = 1 Then
            
        End If
        rsMap.movenext
    Loop
End Function

'判断当前字段是否需要判断重复值的问题
'重复的话返回FALSE
'不重复的话返回TRUE
Private Function JudgeRules_Repeat(ByVal vMappingIndex As Long, ByRef vRsDataExcel As RecordSet, ByRef vRsDataTable As RecordSet) As Boolean
    On Error GoTo IFERR
    Dim szFilter As String
    JudgeRules_Repeat = True
    If A_rsRules.State <> adStateOpen Then
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    If A_rsRules.RecordCount <= 0 Then
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    Dim rsDataTable As RecordSet
    Dim szRuleFieldName As String
    Dim szRuleFieldValue As String
    Dim szFilterString As String
    
    
    szFilter = ""
    szFilter = " B_ID=" & vMappingIndex
    A_rsRules.Filter = szFilter
    If A_rsRules.RecordCount <= 0 Then
        A_rsRules.Filter = ""
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    szFilter = szFilter & " And B_OnlyOne=1"
    'A_rsRules.Filter = " B_OnlyOne=1"
    Debug.Print szFilter
    A_rsRules.Filter = szFilter
    If A_rsRules.RecordCount <= 0 Then
        A_rsRules.Filter = ""
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    
    Dim szFieldCN As String
    Dim szFieldValue As String
    Dim szTip As String
    
    
    Set rsDataTable = vRsDataTable.Clone
    A_rsMappingDetail.Filter = " B_ID=" & vMappingIndex
    Do While Not A_rsRules.EOF
        A_rsMappingDetail.Filter = " B_FieldName='" & Trim$(A_rsRules!B_FieldName) & "'"
        If A_rsMappingDetail.RecordCount > 0 Then
            szRuleFieldName = Trim$(A_rsRules!B_FieldName)
            szRuleFieldValue = Trim(vRsDataExcel(Val(A_rsMappingDetail!B_ExcelColIndex) - 1))
            szFilterString = " " & szRuleFieldName & "='" & szRuleFieldValue & "'"
            
            rsDataTable.Filter = szFilterString
            If rsDataTable.RecordCount > 0 Then
                JudgeRules_Repeat = False
                szFieldCN = Get_FieldCN_110901(Trim$(A_rsRules!B_FieldName))
                szTip = szFieldCN & ":" & szRuleFieldValue & "重复" & vbNewLine
                szTip = szTip & "请记录下本订单号，点击确定要继续导入下一条数据！"
                MsgBox szTip, vbOKOnly + vbInformation, "提示"
                Exit Function
            End If
            
        End If
        A_rsRules.movenext
    Loop
    
    
    Exit Function
IFERR:
    Dim szErr As String
    szErr = Err.Description & vbNewLine
    szErr = szErr & "发生于JudgeRules_Repeat！"
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Function


Private Sub InitRules()
    Set A_rsRules = New RecordSet
    strSQL = "Select * From G_ExcelMappingTableRules"
    A_rsRules.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
End Sub

Private Sub InitMappingDetail()
    InitMappingDetail
    InitRules
    
    
    Set A_rsMappingDetail = New RecordSet
    strSQL = "Select * From G_ExcelMappingTableDetail"
    A_rsMappingDetail.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
End Sub


'获取单表110901中字段的中文名称
Private Function Get_FieldCN_110901(ByVal vFieldName As String) As String
    Dim rs As New RecordSet
    Dim szReturn As String
    
    strSQL = "Select * From dbo.G_BLSField Where B_ObjectID='110901'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        szReturn = ""
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    
    rs.Filter = " B_FieldName='" & vFieldName & "'"
    If rs.RecordCount > 0 Then
        szReturn = rs!B_CnName
    Else
        szReturn = ""
    End If
    rs.Close
    Set rs = Nothing
    
    Get_FieldCN_110901 = szReturn
End Function

'从路径中获取文件名
Public Function Get_FileName_FromPath(ByVal vStrPath As String) As String
    Dim i As Long
    i = InStrRev(vStrPath, "\")
    Get_FileName_FromPath = Mid$(vStrPath, i + 1)
End Function

'拷贝一个文件
Public Function CopyFile(ByVal voPath As String, vtPath As String) As Boolean
    Dim szFileName As String
    Dim szTPathFileName As String
    
    '检测目标路径是否存在， 不存在则创建
    Judge_Folder_Exists_And_AddNew vtPath
    
    szFileName = Get_FileName_FromPath(voPath)
    If Right$(vtPath, 1) = "\" Then
        szTPathFileName = vtPath & szFileName
    Else
        szTPathFileName = vtPath & "\" & szFileName
    End If
    Debug.Print szTPathFileName
    
    FileCopy voPath, szTPathFileName
End Function

'判断一个路径（文件夹）是否存在
Public Function Judge_Folder_Exists(ByVal vPath As String) As Boolean
    If Dir(vPath, vbDirectory) = "" Then
        Judge_Folder_Exists = False
    Else
        Judge_Folder_Exists = True
    End If
End Function

'一个文件路径，从分区根目录开始检测，逐层文件夹检测，路径不存在的话
'则逐层创建
Private Sub Create_Folder(ByVal vPath As String)
    Dim arr01
    Dim szFilterSym As String
    Dim i As Long, J As Long
    Dim szFilePath As String
    
    
    szFilterSym = "\"
    arr01 = Split(vPath, szFilterSym)
    
    
    If UBound(arr01) >= 1 Then
        szFilePath = arr01(0)
        For i = 1 To UBound(arr01)
            szFilePath = szFilePath & "\" & arr01(i)
            If Judge_Folder_Exists(szFilePath) = False Then
                MkDir szFilePath
            End If
        Next
    End If

End Sub

'判断一个路径是否存在，不存在则创建
Public Sub Judge_Folder_Exists_And_AddNew(ByVal vPath As String)
    '把最后一个\去掉（微软认同的格式）
    If Right$(vPath, 1) = "\" Then
        vPath = Left$(vPath, Len(vPath) - 1)
    End If
    
    
    If Judge_Folder_Exists(vPath) = False Then
        Create_Folder vPath
    End If
End Sub

'弹出保存文件的对话框来保存文件
Public Function SaveFile_PopUpDialog() As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = "文本文件 (*.txt)|*.txt|所有文件(*.*)|*.*"
        .ShowSave
        szFileName = .FileName
    End With
    
    SaveFile_PopUpDialog = Trim(szFileName)
    
    Set o = Nothing
End Function

'调用的时候自定义过滤性文件类型
Public Function OpenFile_Dialog_Filter(ByVal vFilterString As String) As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        '.Filter = "Excel文件 (*.xls)|*.xls|所有文件(*.*)|*.*|Excel文件 (*.xlsx)|*.xlsx"
        .Filter = vFilterString
        .ShowOpen
        szFileName = .FileName
    End With
    
    OpenFile_Dialog_Filter = Trim(szFileName)
    
    Set o = Nothing
End Function


Public Sub WriteLog(ByVal vContent As String)
    If Len(Trim$(W_LogFileName)) <= 0 Then
        MsgBox "请先初始化日志文件名称！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    WriteFileContent W_LogFileName, Now & ":"
    WriteFileContent W_LogFileName, vContent & vbNewLine
End Sub


'需要在工程->引用  中引用“Microsoft Scripting Runtime”
'vFolderPath：文件夹路径。查询该文件夹下的所有子文件夹
'vBSubFolders：是否继续无穷尽的枚举子文件夹
Private Sub TraversalFolders(ByVal vFolderPath As String, ByVal vBSubFolders As Boolean)
    Dim fs, F, f1, S, sf
    Dim hs, H, H1, hf
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFolder(vFolderPath)
    
    Set sf = F.SubFolders
    
    For Each f1 In sf
        'List1.AddItem vFolderPath & "\" & f1.name
        If vBSubFolders = True Then
            Call TraversalFolders(vFolderPath & "\" & f1.Name, vBSubFolders)
        End If
    Next
End Sub



'遍历一个文件夹下的所有文件并且返回一个记录集
'该记录集只有一个字段B_Field0包含了文件的绝对路径
Public Function TraversalFilesReturnRS(ByVal vFolderPath As String) As RecordSet
    Dim fs, F, f1, fc, S
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFolder(vFolderPath)
    Set fc = F.Files
    
    
    Dim rs As New RecordSet
    rs.Fields.Append "B_Field0", adVarChar, 1000
    rs.Open
    
    For Each f1 In fc
        rs.AddNew
        rs(0) = vFolderPath & "\" & f1.Name
    Next
    
    Set TraversalFilesReturnRS = rs
End Function

'vPathAndFileName:绝对路径+文件名称（带有后缀名）
'创建一个文件，如果其中路径不存在则创建
'如果该文件已经存在，则不会再创建本文件并且返回FALSE
'如果文件不存在，则创建并且返回TRUE
Public Function CreateFile(ByVal vPathAndFileName As String) As Boolean
    Dim lngHandle As Long '句柄
    Dim cslString1 As New clsString
    
    CreateFile = False
    
    '格式化参数
    vPathAndFileName = cslString1.KillCrlfSpace(vPathAndFileName)
    
    '判断路径是否为空
    If Len(Trim$(vPathAndFileName)) <= 0 Then
        CreateFile = False
        Exit Function
    End If
    
    lngHandle = FreeFile() '取得句柄
    
    Dim szPath As String
    Dim i As Long
    i = InStrRev(vPathAndFileName, "\") '获取最后一个斜线的INDEX
    If Mid$(vPathAndFileName, i - 1, 1) = ":" Then
        '如果是在盘符下的根目录
        szPath = Left(vPathAndFileName, i)
    Else
        '非分区的根目录
        szPath = Left$(vPathAndFileName, i - 1)
    End If
    
    '检测路径是否存在，如果不存在则逐层创建
    Judge_Folder_Exists_And_AddNew szPath
    
    If JudgeFileExist(vPathAndFileName) = False Then
        '创建一个文件
        Dim fso As New FileSystemObject
        Dim fileStream As TextStream
        Set fileStream = fso.CreateTextFile(vPathAndFileName)
        fileStream.Close
        
        Exit Function
    End If

    CreateFile = False
End Function

'在2015年10月8日 13:26:35弃用
'写入之前判断当前文件是否存在，如果不存在则创建
'创建的时候要写入列头，非创建的时候只要写入数据即可。
'xlSheet：从该参数中可以获取第一行（即列头）
'vExcludeCols：排除前多少列
'vCols：要导入的列的数目
'vErrFileName：每次调用导入函数的时候该文件名称都不一样，判断当前文件名是否存在
'vRowIndex：导入的时候出错的行号。会将本行的所有数据写入出错日志表
Private Sub WriteErrExcel(ByRef xlSheet As excel.Worksheet, _
    ByVal vExcludeCols As Long, ByVal vCols As Long, _
    ByVal vErrFileName As String, ByVal vRowIndex As Long)
    
    
    
    Dim clsExcel1 As New clsExcel
    Dim i As Long
    Dim szTemp As String
    
    '创建错误数据记录文件
    If JudgeFileExist(vErrFileName) = False Then
        CreateFile vErrFileName
        
        '创建的同时要写入列头
        For i = (1 + vExcludeCols) To (vCols + vExcludeCols)
            szTemp = xlSheet.Cells(1, i)
            clsExcel1.WriteContentAtCell 1, i, szTemp
        Next
    End If
    
    
    '写入没有导入成功的数据
    For i = (1 + vExcludeCols) To (vCols + vExcludeCols)
         
    Next
End Sub

Private Sub writeLogImportLog(ByVal vContent As String)
    Dim g_szFileName_ImportLog As String
    Dim g_lImportLog As Long
    clsLog.WriteFileContent_WithSwitch g_szFileName_ImportLog, vContent, g_lImportLog
End Sub

'需要将图片显示在PictureBox控件中，才可以成功上传到DB
'读取的时候必须使用PerpertyBag_DownloadFromDB才能正确下载并且显示在PictureBox中
Public Function PerpertyBag_Upload2DB(ByRef oPicture As Object, ByVal vTable As String, _
    ByVal vKeyField As String, ByVal vKeyFieldValue As String, ByVal vPictureField As String) As Boolean

    On Error GoTo IFERR
    Dim strSQL As String
    Dim rs As New RecordSet
    strSQL = "Select * From " & vTable & " where " & vKeyField & " ='" & vKeyFieldValue & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Dim Bag As PropertyBag
    Dim buff() As Byte
    Set Bag = New PropertyBag
    Bag.WriteProperty "Image", oPicture.Picture
    ReDim buff(LenB(Bag.Contents))
    buff = Bag.Contents
    
    rs(Trim$(vPictureField)) = buff
    rs.Update
    Set rs = Nothing
    Set Bag = Nothing
    
    PerpertyBag_Upload2DB = True
    Exit Function
IFERR:
    Dim szErr As String
    PerpertyBag_Upload2DB = False
    szErr = "错误发生于上传中。" & vbNewLine & Err.Description
    'MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Function

'必须配套使用PerpertyBag_Upload2DB，才可以正确下载并且显示图片
Public Function PerpertyBag_DownloadFromDB(ByRef oPicture As Object, ByVal vTable As String, _
    ByVal vKeyField As String, ByVal vKeyFieldValue As String, ByVal vPictureField As String) As Boolean
    
    On Error Resume Next
    Dim strSQL As String
    Dim rs As New RecordSet
    strSQL = "Select * From " & vTable & " where " & vKeyField & " ='" & vKeyFieldValue & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Dim buff() As Byte
    Dim Bag As New PropertyBag
    

    buff = rs.Fields(Trim$(vPictureField)).Value
    
    
    Set Bag = New PropertyBag
    Bag.Contents = buff
    Call Bag.WriteProperty("Image", buff)
    
    '在加载图片前先清空控件内的图片
    oPicture.Picture = LoadPicture("")
    Set oPicture.Picture = Bag.ReadProperty("Image")
    
    
    Set rs = Nothing
    Set Bag = Nothing
    
    PerpertyBag_DownloadFromDB = True
    Exit Function
IFERR:
    Dim szErr As String
    PerpertyBag_DownloadFromDB = False
    szErr = "错误发生于下载中。" & vbNewLine & Err.Description
    'MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Function

Public Sub DownloadPic(ByRef vFld As ADODB.Field, ByVal vPath As String)
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    
    If Not IsNull(vFld) Then
        Stream.Type = adTypeBinary
        'Stream.Mode = adModeReadWrite
        Stream.Open
        Stream.Write vFld
        Stream.SaveToFile vPath, adSaveCreateOverWrite
        Stream.Close
    End If
End Sub

'参数介绍：
'vRs：包含有图片资源的数据源
'vPicField：保存图片文件的字段名
'oCtl：用于显示的控件。PictureBox、Image
Public Sub PicShowToCtl(ByRef vFld As ADODB.Field, ByRef oCtl As Object)
    On Error GoTo IFERR
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    

    oCtl.Picture = LoadPicture("")
    If Not IsNull(vFld) Then
        Stream.Type = adTypeBinary
        Stream.Open
        Stream.Write vFld
        Stream.SaveToFile "filename", adSaveCreateOverWrite
        oCtl.Picture = LoadPicture("filename")
        Stream.Close
    End If
    
    Set Stream = Nothing
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "错误发生于下载图片中，" & Err.Description
    
End Sub


'fld：记录集中的字段
'vFilePath：图片文件的绝对路径，包含图片文件名和扩展名
Public Sub PicSaveToDB(ByRef fld As ADODB.Field, ByVal vFilePath As String)
    Const blocksize = 4096
    Dim bytedata() As Byte
    Dim numblocks As Long
    Dim filelength As Long
    Dim leftover As Long
    Dim sourcefile As Long
    Dim i As Long
    
    '如果没有文件，那么直接退出
    If Len(vFilePath) <= 0 Then
        Exit Sub
    End If
    
    sourcefile = FreeFile
    
    Open Trim(vFilePath) For Binary Access Read As sourcefile
    filelength = LOF(sourcefile)
    
    If filelength = 0 Then
        Close sourcefile
        'MsgBox Trim(vFilePath) & "无内容或不存在！"
    Else
        numblocks = filelength \ blocksize
        leftover = filelength Mod blocksize
        fld.Value = Null
        
        ReDim bytedata(blocksize)
        
        For i = 1 To numblocks
            Get sourcefile, , bytedata()
            fld.AppendChunk bytedata()
        Next
        
        ReDim bytedata(leftover)
        Get sourcefile, , bytedata()
        fld.AppendChunk bytedata()
        Close sourcefile
    End If
End Sub


Public Sub WriteFileContentChoice(ByVal strFileName As String, ByVal strWrite As String, ByVal vChoice As Long)
    Dim lngHandle As Long '句柄
    
    If vChoice = 0 Then
        Exit Sub
    End If
    
    lngHandle = FreeFile() '取得句柄
    

    Open strFileName For Append As lngHandle    '打开文件
    Print #lngHandle, strWrite    '写入文本
    Close lngHandle    '关闭文件
End Sub


Public Function ShowSaveFileDialog(Optional ByVal vFilterString As String = "所有文件(*.*)|*.*") As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = vFilterString
        .ShowSave
        szFileName = .FileName
    End With
    
    ShowSaveFileDialog = Trim(szFileName)
    
    Set o = Nothing
End Function
