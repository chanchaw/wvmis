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
Private A_rsBLRFields As RecordSet   '��Ӧ���̱��������е�����
Private A_rsBLRFormTools As RecordSet
Private A_rsBLSFields As RecordSet  '��Ӧ���������е�����   2017��11��23��

Private A_GroupFields As String
Private A_SumFields As String    '����ֶ�
Private A_AvegFields As String   '��ƽ���ֶ�
Private A_arrGroupFields() As String
Private A_arrSumFields() As String   '����ֶ�
Private A_arrAvegFields() As String   '��ƽ���ֶ�

Private strSQL As String

Private rsField As RecordSet  '��G_BLRField������

'��ʼ�������ֶε�����
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
    
    '��ȡ��ǰѡ����ѡ���еĵ�Ԫ�������
    'A_VsFlexGrid.TextMatrix(A_VsFlexGrid.RowSel,A_VsFlexGrid.ColSel)
    
    '�����ֶ�����ȡ�е�INDEX
    'A_VsFlexGrid.ColIndex ("�ֶ�����")
    
    
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

'������ؼ��������
Public Sub FillGrid(ByRef vRs As RecordSet)
    Dim i As Long
    With A_VsFlexGrid
        Set .DataSource = vRs
        .AllowUserResizing = flexResizeBoth
        .SubtotalPosition = flexSTBelow
        .OutlineBar = flexOutlineBarSimpleLeaf
        
        '�����������иߡ������п�
        '.RowHeight(-1) = 600 '���������и�Ϊ800
        .ColWidth(-1) = 1000  '���������п�Ϊ500
        .RowHeight(-1) = 400
        
        .EditAble = flexEDKbd
        
        
        '���ÿ�����趨�п�
        .AutoSizeMode = flexAutoSizeColWidth
        
        '.BackColorSel = RGB(255, 207, 49)
        .BackColorSel = RGB(0, 128, 255)
        .ColWidth(0) = 300
        .RowHeight(0) = 400
        
        
        '�趨��ͷ��ͼƬ
        Dim szPicturePath As String
        szPicturePath = App.Path & "\image\GridHead.bmp"
        .cell(flexcpPicture, 0, 0, 0, .cols - 1) = LoadPicture(szPicturePath)
        .cell(flexcpPictureAlignment, 0, 0, 0, .cols - 1) = flexPicAlignStretch


        '�����ͷ��Caption
        SetColCaption
'
'        Dim myFont As New StdFont
'        myFont.name = "����"
'        Set .Cell(flexcpFont, 0, 0, .Rows - 1, .cols - 1) = myFont
        
        '���ݿ��������е���������ÿ�еĿ���
        
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        
        .DataMode = flexDMFree
        .ExplorerBar = flexExSortAndMove
        .Redraw = True
    End With
    
    '���鼰���
    GroupAndSum A_VsFlexGrid
    
    '���ÿؼ�
    SetControls
    
    '��ʾ�к�
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
        MsgBox "��δ��ʼ����", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    Set A_rsBLR = New RecordSet
    strSQL = "SELECT * FROM G_BLR WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLR.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set A_rsBLRFields = New RecordSet
    strSQL = "SELECT * FROM G_BLRField WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLRFields.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    '����������������
    Set A_rsBLSFields = New RecordSet
    strSQL = "SELECT * FROM G_BLSField WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLSFields.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    Set A_rsBLRFormTools = New RecordSet
    strSQL = "SELECT * FROM G_BLRFormTools WHERE B_ObjectID='" & A_ObjectID & "'"
    A_rsBLRFormTools.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    '��ʼ�������ֶ�
    A_GroupFields = IIf(IsNull(A_rsBLR!B_GroupFields), "", A_rsBLR!B_GroupFields)
    If Len(A_GroupFields) > 0 Then
        A_arrGroupFields = Split(A_GroupFields, ",")
    End If
    
    
    '��ʼ������ֶ�
    A_SumFields = IIf(IsNull(A_rsBLR!B_SumFields), "", A_rsBLR!B_SumFields)
    If Len(A_SumFields) > 0 Then
        A_arrSumFields = Split(A_SumFields, ",")
    End If
    
    
    '��ʼ����ƽ���ֶ�
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
            '�����ֶ�ѭ��
            For i = 0 To UBound(A_arrGroupFields)
                If Gm.ArrNotEmpty(A_arrSumFields) = True Then
                    For J = 0 To UBound(A_arrSumFields)
                        '.Subtotal flexSTSum, .ColIndex(A_arrGroupFields(i)), .ColIndex(A_arrSumFields(j)), , , RGB(0, 0, 255), True, "%sС�ƣ�"
                        .Subtotal flexSTSum, .ColIndex(A_arrGroupFields(i)), .ColIndex(A_arrSumFields(J)), , , RGB(0, 0, 255), True, "С�ƣ�"
                    Next
                Else
                    '���û�л����ֶΣ���ô�趨��һ��Ϊ���ܣ�
                    '����������ɫ������ؼ��ı�����ɫһ��
                    '�Դﵽû�л��ܵ��Ӿ�Ч��
                    .Subtotal flexSTSum, .ColIndex(A_arrGroupFields(i)), 1, , , .BackColor, False, ""
                End If
                
                '��ƽ���ֶ�
                If Gm.ArrNotEmpty(A_arrAvegFields) = True Then
                    For J = 0 To UBound(A_arrAvegFields)
                        .Subtotal flexSTAverage, .ColIndex(A_arrGroupFields(i)), .ColIndex(A_arrAvegFields(J)), , , RGB(0, 0, 255), True, "%sƽ����"
                    Next
                End If
            Next
        End If
        
        '�����ܼ�
        If Len(A_SumFields) > 0 Then
            '����ֶ�
            If Gm.ArrNotEmpty(A_arrSumFields) = True Then
            For i = 0 To UBound(A_arrSumFields)
                .Subtotal flexSTSum, -1, .ColIndex(A_arrSumFields(i)), , , RGB(255, 0, 0), True, "%s �����ܼƣ�"
            Next
            End If
            
            
            '��ƽ�����ֶ�
            If Gm.ArrNotEmpty(A_arrAvegFields) = True Then
            For i = 0 To UBound(A_arrAvegFields)
                .Subtotal flexSTAverage, -1, .ColIndex(A_arrAvegFields(i)), , , RGB(0, 0, 255), True, "%sƽ����"
            Next
            End If
            
            
            If vVsFlexGrid.Rows > 1 Then
            For i = 1 To vVsFlexGrid.cols - 1
                If vVsFlexGrid.ColHidden(i) = False Then
                    vVsFlexGrid.cell(flexcpText, vVsFlexGrid.Rows - 1, i, vVsFlexGrid.Rows - 1, i) = "�����ܼƣ�"
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

'������ͷ�ĺ���ͬʱ�����п���
'ͬʱ���ö��뷽ʽ
'���ݵ���������
Public Sub SetColCaption_BLS()
    Dim i As Long
    Dim szCaption As String
    Dim szField As String
    Dim szTemp As String
    
    For i = 0 To A_VsFlexGrid.cols - 1
        With A_VsFlexGrid
            szField = .ColKey(i)
            A_rsBLSFields.Filter = " B_FieldName='" & szField & "'"
            Debug.Print "�ֶ� = " & szField
            If A_rsBLSFields.RecordCount > 0 Then
                szCaption = IIf(IsNull(A_rsBLSFields!B_CnName), "", A_rsBLSFields!B_CnName)
                .ColWidth(i) = IIf(IsNull(A_rsBLSFields!B_GridWidth), 1000, A_rsBLSFields!B_GridWidth)
                
                Select Case A_rsBLSFields!B_Alignment
                    Case 0 '����
                        .ColAlignment(i) = flexAlignLeftCenter
                    Case 1 '����
                        .ColAlignment(i) = flexAlignRightCenter
                    Case 2 '����
                        .ColAlignment(i) = flexAlignCenterCenter
                End Select
                
                
                '�е�Format
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
    
        Case 3 '����
            SetColCaption_BLS
        Case 4 '����
        Case 5 '���̱���
            SetColCaption_BLR
    End Select
    
    
            '�趨��ͷ��ͼƬ
    Dim szPicturePath As String
    szPicturePath = App.Path & "\image\GridHead.bmp"
    A_VsFlexGrid.cell(flexcpPicture, 0, 0, 0, A_VsFlexGrid.cols - 1) = LoadPicture(szPicturePath)
    A_VsFlexGrid.cell(flexcpPictureAlignment, 0, 0, 0, A_VsFlexGrid.cols - 1) = flexPicAlignStretch
End Sub


'������ͷ�ĺ���ͬʱ�����п���
'ͬʱ���ö��뷽ʽ
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
                    Case 0 '����
                        .ColAlignment(i) = flexAlignLeftCenter
                    Case 1 '����
                        .ColAlignment(i) = flexAlignRightCenter
                    Case 2 '����
                        .ColAlignment(i) = flexAlignCenterCenter
                End Select
                
                
                '�е�Format
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

'��ӡԤ��
Public Sub ExportExcel(ByVal vFilePath As String)
    Dim i As Long
    For i = 0 To A_VsFlexGrid.cols - 1
        A_VsFlexGrid.ColDataType(i) = flexDTString
    Next

    A_VsFlexGrid.SaveGrid vFilePath, flexFileCustomText, True
End Sub

'Ϊ����ʾFormat
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

'��ȡ��ǰ���µ�ָ���е�ֵ
Public Function GetCurRowSpecColValue(ByVal vFieldName As String) As String
    GetCurRowSpecColValue = A_VsFlexGrid.cell(flexcpText, A_VsFlexGrid.Row, GetColIndex(Trim$(vFieldName)))
End Function

'��ʾ��Ԫ���ڵĿؼ�
Private Sub SetControls()
    Dim strSQL As String
    Dim rs As RecordSet
    Dim ctl
    Dim lColIndex As Long  '�ֶ����ڵ���ColIndex
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
        
    
        '���ÿؼ��Ķ��뷽ʽ
        '��ѭ�ֶεĶ��뷽ʽ
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
        
        
        '�Ƿ�����
        lTemp = IIf(IsNull(rsField!B_Locked), 0, rsField!B_Locked)
        If lTemp = 1 Then
            '���������ɱ༭
            
        End If
        
        rs.movenext
    Loop
    
    
    rs.Close
    Set rs = Nothing
End Sub

'�жϵ�ǰ���Ƿ���Ա༭
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

