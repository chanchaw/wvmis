VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVsFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private A_ObjectID As String
Private WithEvents A_VsFlexGrid As VSFlexGrid
Attribute A_VsFlexGrid.VB_VarHelpID = -1
Private A_Initialized As Boolean

Private A_rsBLR As RecordSet
Private A_rsBLRFields As RecordSet   '对应工程报表网格列的设置
Private A_rsBLRFormTools As RecordSet
Private A_rsBLSFields As RecordSet  '对应单表网格列的设置   2017年11月23日

Private A_GroupFields As String
Private A_SumFields As String    '求和字段
Private A_AvegFields As String   '求平均字段
Private A_arrGroupFields() As String
Private A_arrSumFields() As String   '求和字段
Private A_arrAvegFields() As String   '求平均字段

Private strSQL As String

Private rsField As RecordSet  '表G_BLRField的设置

'初始化关于字段的设置
Private Sub InitRsField()
    Set rsField = New RecordSet
    strSQL = "Select * From G_BLRField where B_ObjectID='" & A_ObjectID & "'"
    rsField.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
End Sub

Public Sub InitCls(ByVal vObjectID As String, ByRef vFlexGrid As VSFlexGrid)
    A_ObjectID = vObjectID
    Set A_VsFlexGrid = vFlexGrid
    
    A_Initialized = True
    
    GetReportPara
    InitRsField
End Sub

Private Sub A_VsFlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'A_VsFlexGrid.cell(flexcpText, Row, Col + 1) = A_VsFlexGrid
    
    '获取当前选中行选中列的单元格的内容
    'A_VsFlexGrid.TextMatrix(A_VsFlexGrid.RowSel,A_VsFlexGrid.ColSel)
    
    '根据字段名获取列的INDEX
    'A_VsFlexGrid.ColIndex ("字段名称")
    
    
End Sub

Private Sub A_VsFlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, _
    Cancel As Boolean)
    
    If EditAble(Col) = False Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub Class_Initialize()
    A_Initialized = False
End Sub

'向网格控件填充数据
Public Sub FillGrid(ByRef vRs As RecordSet)
    Dim i As Long
    With A_VsFlexGrid
        Set .DataSource = vRs
        .AllowUserResizing = flexResizeBoth
        .SubtotalPosition = flexSTBelow
        .OutlineBar = flexOutlineBarSimpleLeaf
        
        '先设置所有行高、所有列宽
        '.RowHeight(-1) = 600 '设置所有行高为800
        .ColWidth(-1) = 1000  '设置所有列宽为500
        .RowHeight(-1) = 400
        
        .EditAble = flexEDKbd
        
        
        '针对每个列设定列宽
        .AutoSizeMode = flexAutoSizeColWidth
        
        '.BackColorSel = RGB(255, 207, 49)
        .BackColorSel = RGB(0, 128, 255)
        .ColWidth(0) = 300
        .RowHeight(0) = 400
        
        
        '设定列头的图片
        Dim szPicturePath As String
        szPicturePath = App.Path & "\image\GridHead.bmp"
        .cell(flexcpPicture, 0, 0, 0, .cols - 1) = LoadPicture(szPicturePath)
        .cell(flexcpPictureAlignment, 0, 0, 0, .cols - 1) = flexPicAlignStretch


        '填充列头的Caption
        SetColCaption
'
'        Dim myFont As New StdFont
'        myFont.name = "宋体"
'        Set .Cell(flexcpFont, 0, 0, .Rows - 1, .cols - 1) = myFont
        
        '根据开发环境中的配置设置每列的宽度
        
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        
        .DataMode = flexDMFree
        .ExplorerBar = flexExSortAndMove
        .Redraw = True
    End With
    
    '分组及求和
    GroupAndSum A_VsFlexGrid
    
    '设置控件
    SetControls
    
    '显示行号
    ShowRowNumber
    
End Sub

Private Sub ShowRowNumber()
    Dim i As Long
    For i = 1 To A_VsFlexGrid.Rows - 1
        A_VsFlexGrid.cell(flexcpText, i, 0, i, 0) = i
    Next
End Sub

Private Sub GetReportPara()
    On Error Resume Next
    If A_Initialized = False Then
        MsgBox "类未初始化！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    Set A_rsBLR = New RecordSet
    strSQL = "SELECT * FROM G_BLR WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLR.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set A_rsBLRFields = New RecordSet
    strSQL = "SELECT * FROM G_BLRField WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLRFields.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    '单表的网格列设置
    Set A_rsBLSFields = New RecordSet
    strSQL = "SELECT * FROM G_BLSField WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLSFields.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    Set A_rsBLRFormTools = New RecordSet
    strSQL = "SELECT * FROM G_BLRFormTools WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLRFormTools.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    '初始化分组字段
    A_GroupFields = IIf(IsNull(A_rsBLR!B_GroupFields), "", A_rsBLR!B_GroupFields)
    If Len(A_GroupFields) > 0 Then
        A_arrGroupFields = Split(A_GroupFields, ",")
    End If
    
    
    '初始化求和字段
    A_SumFields = IIf(IsNull(A_rsBLR!B_SumFields), "", A_rsBLR!B_SumFields)
    If Len(A_SumFields) > 0 Then
        A_arrSumFields = Split(A_SumFields, ",")
    End If
    
    
    '初始化求平均字段
    A_AvegFields = IIf(IsNull(A_rsBLR!B_AverageFields), "", A_rsBLR!B_AverageFields)
    If Len(A_AvegFields) > 0 Then
        A_arrAvegFields = Split(A_AvegFields, ",")
    End If
End Sub

Private Sub GroupAndSum(ByRef vVsFlexGrid As VSFlexGrid)
    On Error Resume Next
    Dim i As Long, J As Long
    With vVsFlexGrid
        If Len(A_GroupFields) > 0 Then
            '分组字段循环
            For i = 0 To UBound(A_arrGroupFields)
                If Gm.ArrNotEmpty(A_arrSumFields) = True Then
                    For J = 0 To UBound(A_arrSumFields)
                        '.Subtotal flexSTSum, .ColIndex(A_arrGroupFields(i)), .ColIndex(A_arrSumFields(j)), , , RGB(0, 0, 255), True, "%s小计："
                        .Subtotal flexSTSum, .ColIndex(A_arrGroupFields(i)), .ColIndex(A_arrSumFields(J)), , , RGB(0, 0, 255), True, "小计："
                    Next
                Else
                    '如果没有汇总字段，那么设定第一列为汇总，
                    '并且字体颜色和网格控件的背景颜色一致
                    '以达到没有汇总的视觉效果
                    .Subtotal flexSTSum, .ColIndex(A_arrGroupFields(i)), 1, , , .BackColor, False, ""
                End If
                
                '求平均字段
                If Gm.ArrNotEmpty(A_arrAvegFields) = True Then
                    For J = 0 To UBound(A_arrAvegFields)
                        .Subtotal flexSTAverage, .ColIndex(A_arrGroupFields(i)), .ColIndex(A_arrAvegFields(J)), , , RGB(0, 0, 255), True, "%s平均："
                    Next
                End If
            Next
        End If
        
        '报表总计
        If Len(A_SumFields) > 0 Then
            '求和字段
            If Gm.ArrNotEmpty(A_arrSumFields) = True Then
            For i = 0 To UBound(A_arrSumFields)
                .Subtotal flexSTSum, -1, .ColIndex(A_arrSumFields(i)), , , RGB(255, 0, 0), True, "%s 报表总计："
            Next
            End If
            
            
            '求平均的字段
            If Gm.ArrNotEmpty(A_arrAvegFields) = True Then
            For i = 0 To UBound(A_arrAvegFields)
                .Subtotal flexSTAverage, -1, .ColIndex(A_arrAvegFields(i)), , , RGB(0, 0, 255), True, "%s平均："
            Next
            End If
            
            
            If vVsFlexGrid.Rows > 1 Then
            For i = 1 To vVsFlexGrid.cols - 1
                If vVsFlexGrid.ColHidden(i) = False Then
                    vVsFlexGrid.cell(flexcpText, vVsFlexGrid.Rows - 1, i, vVsFlexGrid.Rows - 1, i) = "报表总计："
                    Exit For
                End If
            Next
            End If
        End If
    End With
End Sub

Public Sub SetCollapsed(ByRef vVsFlexGrid As VSFlexGrid)
    Dim i As Long
    For i = 0 To vVsFlexGrid.Rows - 1
        vVsFlexGrid.IsCollapsed(i) = flexOutlineSubtotals
    Next
End Sub

Public Sub SetExpanded(ByRef vVsFlexGrid As VSFlexGrid)
    Dim i As Long
    For i = 0 To vVsFlexGrid.Rows - 1
        vVsFlexGrid.IsCollapsed(i) = flexOutlineExpanded
    Next
End Sub

'设置列头的汉字同时设置列宽度
'同时设置对齐方式
'根据单表来设置
Public Sub SetColCaption_BLS()
    Dim i As Long
    Dim szCaption As String
    Dim szField As String
    Dim szTemp As String
    
    For i = 0 To A_VsFlexGrid.cols - 1
        With A_VsFlexGrid
            szField = .ColKey(i)
            A_rsBLSFields.Filter = " B_FieldName='" & szField & "'"
            Debug.Print "字段 = " & szField
            If A_rsBLSFields.RecordCount > 0 Then
                szCaption = IIf(IsNull(A_rsBLSFields!B_CnName), "", A_rsBLSFields!B_CnName)
                .ColWidth(i) = IIf(IsNull(A_rsBLSFields!B_GridWidth), 1000, A_rsBLSFields!B_GridWidth)
                
                Select Case A_rsBLSFields!B_Alignment
                    Case 0 '居左
                        .ColAlignment(i) = flexAlignLeftCenter
                    Case 1 '居右
                        .ColAlignment(i) = flexAlignRightCenter
                    Case 2 '居中
                        .ColAlignment(i) = flexAlignCenterCenter
                End Select
                
                
                '列的Format
                szTemp = IIf(IsNull(A_rsBLSFields!B_FieldFormat), "", A_rsBLSFields!B_FieldFormat)
                If Len(szTemp) > 0 Then
                    .ColFormat(i) = szTemp
                End If
            Else
                szCaption = ""
                If i <> 0 Then
                .ColHidden(i) = True
                End If
            End If
            
            .cell(flexcpText, 0, i, 0, i) = szCaption
        End With
    Next
End Sub

Public Sub SetColCaption()
    Dim cls1 As New clsDevEnviroment
    Dim lObjectType As Long
    lObjectType = cls1.GetObjectTypeInt(A_ObjectID)
    Select Case lObjectType
    
        Case 3 '单表
            SetColCaption_BLS
        Case 4 '单据
        Case 5 '工程报表
            SetColCaption_BLR
    End Select
    
    
            '设定列头的图片
    Dim szPicturePath As String
    szPicturePath = App.Path & "\image\GridHead.bmp"
    A_VsFlexGrid.cell(flexcpPicture, 0, 0, 0, A_VsFlexGrid.cols - 1) = LoadPicture(szPicturePath)
    A_VsFlexGrid.cell(flexcpPictureAlignment, 0, 0, 0, A_VsFlexGrid.cols - 1) = flexPicAlignStretch
End Sub


'设置列头的汉字同时设置列宽度
'同时设置对齐方式
Public Sub SetColCaption_BLR()
    Dim i As Long
    Dim szCaption As String
    Dim szField As String
    Dim szTemp As String
    
    For i = 0 To A_VsFlexGrid.cols - 1
        With A_VsFlexGrid
            szField = .ColKey(i)
            A_rsBLRFields.Filter = " B_FieldName='" & szField & "'"
            If A_rsBLRFields.RecordCount > 0 Then
                szCaption = IIf(IsNull(A_rsBLRFields!B_CnName), "", A_rsBLRFields!B_CnName)
                .ColWidth(i) = IIf(IsNull(A_rsBLRFields!B_GridWidth), 1000, A_rsBLRFields!B_GridWidth)
                
                Select Case A_rsBLRFields!B_Alignment
                    Case 0 '居左
                        .ColAlignment(i) = flexAlignLeftCenter
                    Case 1 '居右
                        .ColAlignment(i) = flexAlignRightCenter
                    Case 2 '居中
                        .ColAlignment(i) = flexAlignCenterCenter
                End Select
                
                
                '列的Format
                szTemp = IIf(IsNull(A_rsBLRFields!B_FieldFormat), "", A_rsBLRFields!B_FieldFormat)
                If Len(szTemp) > 0 Then
                    .ColFormat(i) = szTemp
                End If
            Else
                szCaption = ""
                If i <> 0 Then
                .ColHidden(i) = True
                End If
            End If
            
            .cell(flexcpText, 0, i, 0, i) = szCaption
        End With
    Next
End Sub

'打印预览
Public Sub ExportExcel(ByVal vFilePath As String)
    Dim i As Long
    For i = 0 To A_VsFlexGrid.cols - 1
        A_VsFlexGrid.ColDataType(i) = flexDTString
    Next

    A_VsFlexGrid.SaveGrid vFilePath, flexFileCustomText, True
End Sub

'为列显示Format
Public Function GetColIndex(ByVal vFieldName As String) As Long
    On Error Resume Next
    Dim i As Long
    For i = 0 To A_VsFlexGrid.cols - 1
        If A_VsFlexGrid.ColKey(i) = vFieldName Then
            GetColIndex = i
            Exit Function
        End If
    Next
        
    GetColIndex = -1
End Function

Public Sub SaveColWidth()
    Dim i As Long
    Dim rs As RecordSet
    Dim dWidth As Double
    Dim szKey As String
    
    strSQL = "SELECT * FROM G_BLRField WHERE B_ObjectID='" & A_ObjectID & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    For i = 1 To A_VsFlexGrid.cols - 1
        dWidth = A_VsFlexGrid.ColWidth(i)
        szKey = A_VsFlexGrid.ColKey(i)
        If dWidth > 0 Then
            rs.Filter = " B_FieldName='" & szKey & "'"
            If rs.RecordCount > 0 Then
                rs!B_GridWidth = dWidth
                rs.Update
            End If
        End If
    Next
    
    rs.Close
    Set rs = Nothing
End Sub

'获取当前行下的指定列的值
Public Function GetCurRowSpecColValue(ByVal vFieldName As String) As String
    GetCurRowSpecColValue = A_VsFlexGrid.cell(flexcpText, A_VsFlexGrid.Row, GetColIndex(Trim$(vFieldName)))
End Function

'显示单元格内的控件
Private Sub SetControls()
    Dim strSQL As String
    Dim rs As RecordSet
    Dim ctl
    Dim lColIndex As Long  '字段所在的列ColIndex
    Dim lTemp As Long
    
    Set rs = New RecordSet
    strSQL = "Select * From G_SGGridControls Where B_ObjectID='" & A_ObjectID & "' Order by B_ID"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        
        Exit Sub
    End If
    
    
    lColIndex = GetColIndex(Trim(rs("B_FieldName")))
    
    Do While Not rs.EOF
        Select Case rs("B_ControlName")
            Case "CheckBox"
                'A_VsFlexGrid.Cell(flexcpChecked, 1, lColIndex, A_VsFlexGrid.Rows - 1, lColIndex)=
                A_VsFlexGrid.ColDataType(lColIndex) = flexDTBoolean

            Case "DropDown"
                
            Case "DropList"
                
            Case "General"
                
            Case "TextBox"
                
                
            Case Else
            
        End Select
        
    
        '设置控件的对齐方式
        '遵循字段的对齐方式
        rsField.Filter = " B_FieldName='" & rs!B_FieldName & "'"
        If rsField.RecordCount > 0 Then
            Select Case rsField!B_Alignment
                Case 0
                    'A_VsFlexGrid.Cell(flexcpAlignment,1,lColIndex,A_VsFlexGrid.Rows-1,lColIndex)=
                    'SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignLeftCenter
                Case 1
                    'SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignRightCenter
                Case 2
                    'SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignCenterCenter
                Case Else
                    'SGrid1.Columns(Trim(rs("B_FieldName"))).Style.PictureAlignment = sgPicAlignCenterCenter
            End Select
        End If
        
        
        '是否锁定
        lTemp = IIf(IsNull(rsField!B_Locked), 0, rsField!B_Locked)
        If lTemp = 1 Then
            '锁定，不可编辑
            
        End If
        
        rs.movenext
    Loop
    
    
    rs.Close
    Set rs = Nothing
End Sub

'判断当前列是否可以编辑
Private Function EditAble(ByVal vColIndex As Long) As Boolean
    Dim szFieldName As String
    Dim lLocked As Long
    szFieldName = A_VsFlexGrid.ColKey(vColIndex)
    
    EditAble = False
    
    rsField.Filter = " B_FieldName='" & szFieldName & "'"
    If rsField.RecordCount <= 0 Then
        Exit Function
    End If
    
    lLocked = IIf(IsNull(rsField!B_Locked), 0, rsField!B_Locked)
    lLocked = Abs(lLocked)
    If lLocked = 1 Then
        EditAble = False
    Else
        EditAble = True
    End If
End Function


Public Sub Dispose()
    A_rsBLR.Close
    Set A_rsBLR = Nothing
    
    A_rsBLRFields.Close
    Set A_rsBLRFields = Nothing
    
    
    
    A_rsBLRFormTools.Close
    Set A_rsBLRFormTools = Nothing
    
    rsField.Close
    Set rsField = Nothing
End Sub


