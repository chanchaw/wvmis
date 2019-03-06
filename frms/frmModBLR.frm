VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{B45CF12E-4E4F-487D-8096-DB3BFE63F435}#1.0#0"; "sg20ou.ocx"
Begin VB.Form frmModBLR 
   BackColor       =   &H00CEDFDE&
   Caption         =   "报表"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "frmModBLR.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   8670
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6405
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8670
      _cx             =   15293
      _cy             =   11298
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   800
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmModBLR.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
         Height          =   5130
         Left            =   0
         TabIndex        =   1
         Top             =   420
         Width           =   8670
         _LayoutVersion  =   1
         _ExtentX        =   15293
         _ExtentY        =   9049
         _DataPath       =   ""
         Bands           =   "frmModBLR.frx":03E2
         Begin VB.PictureBox PctBack 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   480
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   3
            Top             =   180
            Visible         =   0   'False
            Width           =   255
            Begin VB.CommandButton btPopUpWindow 
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Marlett"
                  Size            =   9
                  Charset         =   2
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   4
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin DDSharpGridOLEDB2U.SGGrid SGGrid1 
            Bindings        =   "frmModBLR.frx":05AA
            Height          =   4410
            Left            =   480
            TabIndex        =   2
            Top             =   420
            Width           =   7575
            _cx             =   13361
            _cy             =   7779
            DataMember      =   ""
            DataMode        =   1
            AutoFields      =   -1  'True
            Enabled         =   -1  'True
            GridBorderStyle =   1
            ScrollBars      =   3
            FlatScrollBars  =   1
            ScrollBarTrack  =   0   'False
            DataRowCount    =   0
            BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataColCount    =   2
            HeadingRowCount =   1
            HeadingColCount =   1
            TextAlignment   =   5
            WordWrap        =   -1  'True
            Ellipsis        =   1
            HeadingBackColor=   -2147483633
            HeadingForeColor=   -2147483630
            HeadingTextAlignment=   0
            HeadingWordWrap =   0   'False
            HeadingEllipsis =   1
            GridLines       =   1
            HeadingGridLines=   2
            GridLinesColor  =   -2147483633
            HeadingGridLinesColor=   -2147483632
            EvenOddStyle    =   0
            ColorEven       =   -2147483628
            ColorOdd        =   -2147483624
            UserResizeAnimate=   1
            UserResizing    =   3
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            UserDragging    =   2
            UserHiding      =   2
            CellPadding     =   15
            CellBkgStyle    =   1
            CellBackColor   =   16252927
            CellForeColor   =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   1
            FocusRectColor  =   0
            FocusRectLineWidth=   1
            TabKeyBehavior  =   0
            EnterKeyBehavior=   0
            NavigationWrapMode=   1
            SkipReadOnly    =   0   'False
            DefaultColWidth =   1200
            DefaultRowHeight=   320
            CellsBorderColor=   0
            CellsBorderVisible=   -1  'True
            RowNumbering    =   0   'False
            EqualRowHeight  =   0   'False
            EqualColWidth   =   0   'False
            HScrollHeight   =   0
            VScrollWidth    =   0
            Appearance      =   0
            FitLastColumn   =   0   'False
            SelectionMode   =   2
            MultiSelect     =   2
            AllowAddNew     =   0   'False
            AllowDelete     =   0   'False
            AllowEdit       =   0   'False
            ScrollBarTips   =   0
            CellTips        =   0
            CellTipsDelay   =   1000
            SpecialMode     =   0
            OutlineLines    =   1
            CacheAllRecords =   -1  'True
            ColumnClickSort =   0   'False
            PreviewPaneType =   0
            PreviewPanePosition=   2
            PreviewPaneSize =   2000
            GroupIndentation=   225
            InactiveSelection=   1
            AutoScroll      =   -1  'True
            AutoResize      =   1
            AutoResizeHeadings=   -1  'True
            OLEDragMode     =   0
            OLEDropMode     =   0
            MaxRows         =   4194304
            MaxColumns      =   8192
            NewRowPos       =   1
            CustomBkgDraw   =   0
            AutoGroup       =   -1  'True
            GroupByBoxVisible=   0   'False
            AlphaBlendEnabled=   0   'False
            DragAlphaLevel  =   206
            AutoSearch      =   0
            AutoSearchDelay =   2000
            Format          =   "frmModBLR.frx":05BF
            Caption         =   "frmModBLR.frx":05F1
            ScrollTipColumn =   "frmModBLR.frx":0615
            GroupByBoxText  =   "frmModBLR.frx":0639
            StylesCollection=   "frmModBLR.frx":06BF
            ColumnsCollection=   "frmModBLR.frx":4267
            ValueItems      =   "frmModBLR.frx":5F51
         End
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar2 
         Height          =   420
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8670
         _LayoutVersion  =   1
         _ExtentX        =   15293
         _ExtentY        =   741
         _DataPath       =   ""
         Bands           =   "frmModBLR.frx":68EB
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   3900
            Top             =   5640
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
      Begin VB.Label 状态 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   0
         TabIndex        =   6
         Top             =   5550
         Width           =   8670
      End
   End
End
Attribute VB_Name = "frmModBLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mvarObjectID As String
Public aValue
Private clsReport1 As New clsReport
Private clsCtlShow1 As New clsCtlShow
Private clsSGridShow1 As New clsSGridShow
Private m_ReportObjectID As String

Private m_GroupFields As String
Private m_SumFields As String
Private m_SearchString As String
Private m_NextObject As String

Private iKeyIndex As Integer
Private iCommandIndex As Integer


'==============================
Public fatherFrm As Object
Private mvarfObjectID As String
Private mvarfFieldName As String
Private mvarSendIndex As Integer

Private mvarBillOrDetail As Integer '0 为主表  1为明细表
'==============================


'2012-7-31报表打印权限设置的参数
'==================================
Private m_IsPrintObject As Boolean
Private m_FieldNamePermission As String '数据表中设置可打印权限的字段名称
Private m_RoleBill As String  '主表的角色名称
Private m_DetailName As String  '明细表名称
'==================================

Private A_rsCloneCellColor As RecordSet
Private A_rsObjectPara As New RecordSet

Private strSQL As String

Public Property Let BillOrDetail(ByVal vData As Integer)
    mvarBillOrDetail = vData
End Property

Public Property Get BillOrDetail() As Integer
    BillOrDetail = mvarBillOrDetail
End Property


Public Property Let fObjectID(ByVal vData As String)
    mvarfObjectID = vData
End Property

Public Property Get fObjectID() As String
    fObjectID = mvarfObjectID
End Property

Public Property Let fFieldName(ByVal vData As String)
    mvarfFieldName = vData
End Property

Public Property Get fFieldName() As String
    fFieldName = mvarfFieldName
End Property


Public Property Let SendIndex(ByVal vData As String)
    mvarSendIndex = vData
End Property

Public Property Get SendIndex() As String
    SendIndex = mvarSendIndex
End Property

'==============================



Public Property Let ObjectID(ByVal vData As String)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ObjectID
    ObjectID = mvarObjectID
End Property


Private Sub ActiveBar2_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "导出Excel"
            ExportToExcelA
        Case "刷新"
            PriviewGrid
        Case "打印"
            'PrintPreview
            PrintPreviewEx01
        Case "收缩"
            SGGrid1.CollapseAll
        Case "展开"
            SGGrid1.ExpandAll
        Case "筛选"
            FilterForm
        Case "关闭"
            Unload Me
        Case "删除产量(慎重)"
            DeleteQty
            
        Case "生成成品退仓回修"
            Create120009
            
        Case "设置为无效库存"
            SetInventoryVoid
            
        Case "保存样式"
            
    End Select
End Sub

'生成成品退仓回修
Private Sub Create120009()
    Dim cls1 As New clsBL
    Dim lID As Long
    
    lID = SGGrid1.Rows.Current.Cells(SGGrid1.Columns("B_ID").Position).Value
    cls1.CreateCPTC lID
End Sub

Private Sub FilterForm()
    Dim frm1 As New frmUCTSearch
    With frm1
        Set .rs = Adodc1.RecordSet
        .ObjectID = mvarObjectID
        .FieldType = 5
        .Show vbModal
    End With
    If frm1.OK = True Then
        Adodc1.RecordSet.Filter = ""
        Adodc1.RecordSet.Filter = frm1.strResult
        Debug.Print Adodc1.RecordSet.RecordCount
        ShowSGGrid
    End If
    Unload frm1
    Set frm1 = Nothing
End Sub

Private Sub PrintPreview()
    clsReport1.PrintPreview
End Sub

Private Sub ExportToExcel()
    On Error Resume Next
    Dim sExcelFile As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .DialogTitle = "选择文件"
        .Filter = "EXCEL文件 (*.xls)|*.xls"
        .CancelError = True
        .ShowSave
        sExcelFile = .FileName
    End With
    If Len(sExcelFile) > 2 Then
        SGGrid1.ExportData sExcelFile, sgFormatExcel, sgExportOverwrite + sgExportFieldNames + sgExportRowsOnly
    End If
    Set o = Nothing
End Sub

Private Sub Form_Load()
    ActiveBar21.ClientAreaControl = SGGrid1
    ActiveBar21.RecalcLayout

    GetObjectParameter
    AnimateForm Me

    InitCtl
    
    
    CreatePopUpButton
    
    InitDateTime
End Sub

Public Sub LoadObject()
    LoadCtlParameter
    PriviewGrid
    
    
    '有最大化导航的时候才用到
    Me.WindowState = 2
End Sub

'取得Ctl参数
Private Sub LoadCtlParameter()
    'On Error GoTo IFERR
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim i As Long
    Dim o As Object
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLRFormTools Where B_ObjectID='" & mvarObjectID & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    i = 0
    Do While Not rs.EOF
        For Each o In Me.Controls
            If o.name = rs("B_CtlName") Then
                If i <= UBound(aValue) Then
                    o.Text = aValue(i)
                    i = i + 1
                Else
                    Exit Sub
                End If
            End If
        Next
        rs.movenext
    Loop
    
    Set o = Nothing
    rs.Close
    Set rs = Nothing
    Exit Sub
IFERR:
    Set rs = Nothing
End Sub

Private Sub PriviewGrid()
    On Error GoTo IFERR
    With clsReport1
        .InitClass Me, mvarObjectID, m_ReportObjectID
        .Refresh
    End With
    
    With Adodc1
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .CommandType = adCmdText
        
        
        .ConnectionString = Gm.cnnTool.cnnStr
'        .ConnectionString = Gm.cnnTool.cnn
        .CommandType = adCmdText
        .CommandTimeout = 0
        .RecordSource = clsReport1.sql
        Debug.Print .RecordSource
        .Refresh
    End With
    ShowSGGrid
    GetKeyIndex
    GetCommandIndex
    
    
    '合并单元格
    SetSGGridMergeCells
    
    If IIf(IsNull(A_rsObjectPara!B_Expand), 0, A_rsObjectPara!B_Expand) = 0 Then
        SGGrid1.CollapseAll
    Else
        SGGrid1.ExpandAll
    End If
    
    '当没有记录被查询出来时候要有提示
'    If Adodc1.RecordSet.RecordCount <= 0 Then
'        MsgBox "无符合条件的数据！", vbOKOnly + vbInformation, "提示"
'    End If


    Exit Sub
IFERR:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub ShowSGGrid()
    Set clsSGridShow1 = New clsSGridShow
    With clsSGridShow1
        .ObjectID = mvarObjectID
        .InitClass SGGrid1, 5
        .FillGrid Adodc1.RecordSet
        .ShowGridFormat
        .GroupGrid m_GroupFields, m_SumFields
        .SumGrid ActiveBar21, Adodc1.RecordSet, m_SumFields
    End With
End Sub

Private Sub InitCtl()
    With clsCtlShow1
        .ObjectID = mvarObjectID
        .InitClass ActiveBar21, 3
        .Refresh
    End With
End Sub

'取得参数
Private Sub GetObjectParameter()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLR Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    m_ReportObjectID = rs("B_ReportObject")
    
    Me.width = rs("B_Width")
    Me.height = rs("B_Height")
    Me.Caption = rs("B_ReportName")
    m_NextObject = rs("B_NextObject")
    
    
    m_GroupFields = IIf(IsNull(rs("B_GroupFields")), 0, rs("B_GroupFields"))
    m_SumFields = IIf(IsNull(rs("B_SumFields")), 0, rs("B_SumFields"))
    
    
    Set A_rsObjectPara = rs.Clone
    
    
    rs.Close
    Set rs = Nothing
End Sub


Public Sub Change(ByVal sCtl As String, ByVal sCommand As String)
    'PriviewGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Gm.CacheFrms.DelFrm mvarObjectID
End Sub

Private Sub SGGrid1_DblClick()
    ExecCommand
End Sub

Private Sub ExecCommand()
    On Error GoTo IFERR
    Dim m_Command As String
    Dim aObject
    
    '当是双击组头或者组脚则不执行，直接退出函数
    If SGGrid1.Rows.Current.Type = sgGroupHeader Or SGGrid1.Rows.Current.Type = sgGroupFooter Then
        Exit Sub
    End If
    
    
    '2011-6-20双击报表实现:
    '1.被双击的明细记录的首个字段的值,作为即将打开的子报表的某一个参数
    '2.参照数据表G_CJReportDoublePara设置这个参数的位置
    Dim strSQL As String, i As Long
    Dim rs As RecordSet
    Dim m_aPara
    
    
    
    
    If Len(Trim(m_NextObject)) > 5 Then
        '打开下级对象
        aObject = Split(m_NextObject, ",")
        Set rs = New RecordSet
        strSQL = "Select * From G_CJReportDoublePara Where B_ObjectID='" & mvarObjectID & "'"
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        If rs.RecordCount > 0 Then
            m_Command = clsReport1.GetValueParameter
            Debug.Print m_Command
            m_aPara = Split(m_Command, ",")
            m_Command = ""
            For i = 0 To UBound(m_aPara)
                If i = rs("B_ParaIndex") - 1 Then
                    m_Command = m_Command & SGGrid1.Rows.Current.Cells(iKeyIndex + 1).Text & ","
                    Debug.Print SGGrid1.Rows.Current.Cells(iKeyIndex + 1).Text
                Else
                    m_Command = m_Command & m_aPara(i) & ","
                End If
            Next
            
            m_Command = Left(m_Command, Len(m_Command) - 1)
            Debug.Print m_Command
        Else
            m_Command = clsReport1.GetValueParameter & "," & SGGrid1.Rows.Current.Cells(iKeyIndex + 1).Text
        End If
        
        Gm.Authority.Execute aObject(0), aObject(1), "LoadObject", Nothing, m_Command
        'clsCommand1.Execute aObject(0), aObject(1), "LoadObject", Nothing, m_Command
        
        
        rs.Close
        Set rs = Nothing
    Else
        Dim aCommand
        m_Command = SGGrid1.Rows.Current.Cells(iCommandIndex + 1).Text
        aCommand = Split(m_Command, ",")
        If UBound(aCommand) > 3 Then
            Gm.Authority.Execute aCommand(0), aCommand(1), aCommand(2), Nothing, aCommand(4)
            'clsCommand1.Execute aCommand(0), aCommand(1), aCommand(2), Nothing, aCommand(4)
        End If
    End If
    Exit Sub
IFERR:
    Exit Sub
End Sub

Private Sub GetKeyIndex()
    'iKeyIndex
    Dim i As Integer
    For i = 0 To Adodc1.RecordSet.Fields.Count - 1
        If Adodc1.RecordSet.Fields(i).Properties.item(4).Value = True Then
            iKeyIndex = i
            Exit Sub
        End If
    Next
End Sub

Private Sub GetCommandIndex()
    'iKeyIndex
    Dim i As Integer
    For i = 0 To Adodc1.RecordSet.Fields.Count - 1
        If Adodc1.RecordSet.Fields(i).name = "Command" Then
            iCommandIndex = i
            Exit Sub
        End If
    Next
End Sub

Private Sub ExportToExcelA()
    On Error Resume Next
    Dim Irow, Icol As Integer
    Dim Irowcount, Icolcount As Integer
    Dim Fieldlen() '存字段长度值
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet

'    Dim xlApp As Object
'    Dim xlBook As Object
'    Dim xlSheet As Object
    Dim sgCurRow As SGRow
    
    Dim i As Long
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.add
    Set xlSheet = xlBook.Worksheets(1)

    
    Irowcount = SGGrid1.Rows.Count
    
    Icolcount = 0
    For Icolcount = 0 To SGGrid1.Columns.Count - 1
        If SGGrid1.Columns(Icol).Visible = True Then
            Icolcount = Icolcount + 1
        End If
    Next

    ReDim Fieldlen(Icolcount)
    xlApp.Visible = True '显示表格
    
'    With SGGrid1
'        For iRow = 0 To Irowcount - 1
'            .Row = iRow
'            Set sgCurRow = SGGrid1.Rows.Current
'            If sgCurRow.Type <> sgGroupFooter And sgCurRow.Type <> sgGroupHeader And sgCurRow.Type = sgSimpleRow Then
'
'            For Icol = 0 To Icolcount - 1
'
'                .Col = Icol
'                If .Columns(Icol).Visible = True Then
'                    '添加下面一句后数字列就会在EXCEL中显示为字符串列
'                    '即不可统计
'                    'xlSheet.Cells(Irow + 1, Icol + 1).NumberFormatLocal = "@"
'                    xlSheet.Cells(iRow + 1, Icol + 1).Value = .CellAt(iRow, Icol).Text
'                End If
'
'            Next
'
'            End If
'        Next
'    End With
    
    Irow = 0
    For Each sgCurRow In SGGrid1.Rows
        If sgCurRow.Type <> sgGroupFooter And sgCurRow.Type <> sgGroupHeader Then
            Irow = Irow + 1
            For i = 0 To SGGrid1.Columns.Count
                If SGGrid1.Columns(Icol).Hidden = True Then
                    'MsgBox "是隐藏的列"
                Else
                    xlSheet.Cells(Irow, i).Value = sgCurRow.Cells(SGGrid1.Columns(i).Position).Text
                End If
            Next
        End If
    Next
    
    
    xlApp.Visible = True '显示表格
    'xlBook.Save '保存"
    Set xlApp = Nothing '交还控制给Excel
    Exit Sub
IFERR:
    MsgBox "Excel导出时不正确!", vbExclamation, "Excel"
    Exit Sub

End Sub

'初始化时间
Private Sub InitDateTime()
    On Error Resume Next
    Dim oControl As Object
    Dim rs As RecordSet
    Dim strSQL As String
    
    
    Set rs = New RecordSet
    strSQL = "Select * From G_CJDefaultTime Where B_GroupName='报表默认时间'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    '没有设置参数的话,直接退出本函数
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    
    For Each oControl In Me.Controls
        Select Case oControl.name
            Case "B_SDate"
                If Len(IIf(IsNull(rs("B_SDate")), "", rs("B_SDate"))) > 0 Then
                    oControl.Text = Format(rs("B_SDate"), "YYYY-MM-DD")
                Else
                    oControl.Text = Format(Now, "YYYY-MM-DD")
                End If
                
            Case "B_STime"
                If Len(IIf(IsNull(rs("B_STime")), "", rs("B_STime"))) > 0 Then
                    oControl.Value = Format(rs("B_STime"), "HH:MM:SS")
                Else
                    oControl.Value = Format(Now, "HH:MM:SS")
                End If
                
            Case "B_EDate"
                If Len(IIf(IsNull(rs("B_EDate")), "", rs("B_EDate"))) > 0 Then
                    oControl.Text = Format(rs("B_EDate"), "YYYY-MM-DD")
                Else
                    oControl.Text = Format(Now, "YYYY-MM-DD")
                End If
                
            Case "B_ETime"
                If Len(IIf(IsNull(rs("B_ETime")), "", rs("B_ETime"))) > 0 Then
                    oControl.Value = Format(rs("B_ETime"), "HH:MM:SS")
                Else
                    oControl.Value = Format(Now, "HH:MM:SS")
                End If
                
        End Select
    Next
    
    
    rs.Close
    Set rs = Nothing
    
End Sub


Private Sub SGGrid1_FetchCellStyle(ByVal RowKey As Long, ByVal ColIndex As Long, ByVal CellValue As Variant, ByVal CellStyle As DDSharpGridOLEDB2U.IsgStyle)
    On Error Resume Next
    If Len(Trim(clsSGridShow1.A_FieldsCellColor)) <= 0 Then
        Exit Sub
    End If
    
    
    
    StringToRecordset clsSGridShow1.A_rsCellColor("B_KeyValue"), ",", A_rsCloneCellColor
    
    
    Dim m_str As String
    Dim m_ColIndex As Long
    
    
    m_ColIndex = SGGrid1.Columns(Trim(clsSGridShow1.A_rsCellColor("B_KeyFieldName"))).Position
    m_str = Trim(str(SGGrid1.Rows(RowKey).Cells(SGGrid1.Columns(m_ColIndex).Position).Value))
    
    
    
    '带分割符的字符串转换成的记录集
    A_rsCloneCellColor.MoveFirst
    Do While Not A_rsCloneCellColor.EOF
        If A_rsCloneCellColor(0) = m_str Then
            If clsSGridShow1.A_rsCellColor("B_EffectFields") = "*" Then
                SGGrid1.Rows(RowKey).Style.BackColor = clsSGridShow1.A_rsCellColor("B_Color")
            Else
                m_ColIndex = GetColumnIndexByFieldName(SGGrid1, SGGrid1.Columns(ColIndex).Key)
                SGGrid1.Rows(RowKey).Cells(m_ColIndex).Style.BackColor = clsSGridShow1.A_rsCellColor("B_Color")
            End If
        End If
        A_rsCloneCellColor.movenext
    Loop
    
End Sub

'根据字段名称(存储过程中的字段名)获取当前列的列数
Public Function GetColumnIndexByFieldName(ByRef oSGrid As Object, ByVal m_FieldName As String) As Long
    Dim SGCol As SGColumn
    Dim i As Long
    
    
    For Each SGCol In oSGrid.Columns
        If SGCol.Key = m_FieldName Then
            i = SGCol.ColIndex
            Exit For
        End If
    Next
    GetColumnIndexByFieldName = i
End Function

Private Sub SGGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 2
            '弹出删除产量的菜单
            'ActiveBar2.Bands("Band2").PopupMenu
            
            PopUpRightMenu
    End Select
End Sub


'选择性弹出右键菜单
Private Sub PopUpRightMenu()
    On Error Resume Next
    Dim strSQL As String
    Dim rs As RecordSet
    Dim szPopUpMenuName As String
    
    '1. 获取设置的右键菜单
    Set rs = New RecordSet
    strSQL = "Select * From G_PopUpMenuOnBLR Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    szPopUpMenuName = rs("B_BandName")
    
    rs.Close
    Set rs = Nothing
    
    
    
    '2. 弹出菜单
    If Len(Trim$(szPopUpMenuName)) <= 0 Then
        Exit Sub
    End If
    
    ActiveBar2.Bands(szPopUpMenuName).PopupMenu
End Sub


'删除产量
Private Sub DeleteQty()
    Dim m_BarCode13 As String     '当前删除的条码
    Dim m_ProcessName As String   '当前删除的工序
    Dim m_People As String        '当前删除人员的姓名
    Dim m_QtyAll As Double        '实际重量(整车)
    Dim m_PeopleNumber As Long    '删除后还有几个人
    Dim clsSGridShow1 As New clsSGridShow
    
    Dim strSQL As String
    Dim m_str As String
    Dim i As Long
    
    Select Case mvarObjectID
        Case "130031"
            m_BarCode13 = SGGrid1.Rows.Current.Cells(clsSGridShow1.GetColumnIndexByFieldName(SGGrid1, "B_BarCode13"))
            m_ProcessName = SGGrid1.Rows.Current.Cells(clsSGridShow1.GetColumnIndexByFieldName(SGGrid1, "B_ProcessName"))
            m_People = SGGrid1.Rows.Current.Cells(clsSGridShow1.GetColumnIndexByFieldName(SGGrid1, "B_People"))
    End Select


    If Len(m_BarCode13) <= 0 And Len(m_ProcessName) <= 0 And Len(m_People) <= 0 Then
        MsgBox "当前不可确定产量明细,删除失败!", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
        
    m_str = "您确定要删除" & vbNewLine & "条码为:" & m_BarCode13 & vbNewLine & "工序为:" & m_ProcessName & vbNewLine & "人员为:" & m_People & vbNewLine & "的产量么?"
    
    
    If MsgBox(m_str, vbExclamation + vbYesNo + vbDefaultButton2, "警告") = vbNo Then
        Exit Sub
    End If


    strSQL = " Delete G_CJFlowBillDetail"
    strSQL = strSQL & " From G_CJFlowBill, G_CJFlowBillDetail"
    strSQL = strSQL & " Where G_CJFlowBill.B_ID = G_CJFlowBillDetail.B_ID"
    strSQL = strSQL & " And G_CJFlowBill.B_BarCode13='" & m_BarCode13 & "'"
    strSQL = strSQL & " And G_CJFlowBillDetail.B_ProcessName='" & m_ProcessName & "'"
    strSQL = strSQL & " And G_CJFlowBillDetail.B_People='" & m_People & "'"
    Gm.cnnTool.cnn.Execute strSQL
    
    
    
    
    '刷新网格数据
    PriviewGrid
End Sub



'以下为2012-2-12添加代码:
'自动填充弹出窗体的按钮
'======================================
'生成按钮
Private Sub CreatePopUpButton()
    Dim strSQL As String
    Dim rs As RecordSet
    Dim rs1 As RecordSet
    Dim i As Long
    Dim szTemp As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_PopUpWindowSet Where B_fObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    'GetLeft
    
    
    i = 0
    Do While Not rs.EOF
        i = i + 1
        Load btPopUpWindow(i)
        Load PctBack(i)
        
        PctBack(i).Visible = True
        btPopUpWindow(i).Visible = True
        
        
        Set rs1 = New RecordSet
        szTemp = IIf(IsNull(rs("B_PositionField")), "", rs("B_PositionField"))
        If Len(szTemp) <= 0 Then
            MsgBox "该对象被设置有弹出窗体的查询接口没有设置按钮所在位置！", vbOKOnly + vbInformation, "提示"
            Exit Sub
        End If
        
        'strSQL = "Select * From G_BLRFormTools Where B_FieldName='" & rs("B_fControlName") & "' And B_ObjectID='" & mvarObjectID & "'"
        strSQL = "Select * From G_BLRFormTools Where B_FieldName='" & szTemp & "' And B_ObjectID='" & mvarObjectID & "'"
        Debug.Print strSQL
        rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        
        SetParent PctBack(i).hwnd, ActiveBar21.Bands("栏目" & Trim(rs1("B_BandIndex"))).Tools(Trim(rs1("B_FieldName"))).hwnd
        SetParent btPopUpWindow(i).hwnd, PctBack(i).hwnd

        btPopUpWindow(i).Top = 0
        btPopUpWindow(i).Left = 0
        
        PctBack(i).Top = 20
        'PctBack(i).left = Me.Controls(rs("B_fControlName")).left + Me.Controls(rs("B_fControlName")).Width - btPopUpWindow(0).Width
        PctBack(i).Left = Me.Controls(rs("B_PositionField")).Left + Me.Controls(rs("B_PositionField")).width - btPopUpWindow(0).width
    
    
        PctBack(i).width = PctBack(0).width
        PctBack(i).height = PctBack(0).height
        btPopUpWindow(i).height = btPopUpWindow(0).height
        btPopUpWindow(i).width = btPopUpWindow(0).width
        
        btPopUpWindow(i).Tag = rs("B_tObjectID") & "," & rs("B_fObjectID") & "," & rs("B_fControlName") & "," & rs("B_SendIndex")
        
        
        rs1.Close
        Set rs1 = Nothing
        
            
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub


Private Sub btPopUpWindow_Click(Index As Integer)
    Dim clsCommand1 As New clsCommand
    Dim m_Array
    
    m_Array = Split(btPopUpWindow(Index).Tag, ",")
    
    
    clsCommand1.InitClass
    '数据表G_PopUpWindowSet  B_fObjectID  B_fControlName B_SendIndex  B_tObjectID
    clsCommand1.ExecutePopUp 0, Me, m_Array(1), m_Array(2), m_Array(3), m_Array(0), " ", "LoadObject", Nothing, ""
End Sub


Private Sub ChangeBackColor()
    On Error GoTo IFERR
    Dim oRow As SGRow
    For Each oRow In SGGrid1.Rows
        If oRow.Type = sgSimpleRow And Not oRow.Heading Then
            If Val(oRow.Cells(SGGrid1.Columns("B_DateDiff").Position).Text) > 400 And Val(oRow.Cells(SGGrid1.Columns("B_DateDiff").Position).Text) < 1000 Then
                oRow.Style.BackColor = vbGreen
            End If
       
            If Val(oRow.Cells(SGGrid1.Columns("B_DateDiff").Position).Text) > 1000 Then
                oRow.Style.BackColor = vbRed
            End If
        End If
    Next
    SGGrid1.Redraw
    Exit Sub
IFERR:
    Dim strERR As String
    strERR = Err.Description
    MsgBox strERR, vbOKOnly + vbInformation, "提示"
End Sub
  
  
  
'将带有分割符号的字符串转移到记录集中
Private Sub StringToRecordset(ByVal m_str As String, ByVal m_Filter As String, ByRef rs As RecordSet)
    Dim i As Long
    Dim m_Array
    
    
    m_Array = Split(m_str, m_Filter)
    
    Set rs = New RecordSet
    rs.Fields.Append "B_Field1", adVarChar, 100, adFldIsNullable
    rs.Open
    
    
    For i = 0 To UBound(m_Array)
        If Len(Trim(m_Array(i))) > 0 Then
        
            rs.AddNew
            rs("B_Field1") = m_Array(i)
            rs.Update
        End If
    Next
    
End Sub
   
   
   
   '设置SGGrid是否合并CELLS
Private Sub SetSGGridMergeCells()
    On Error GoTo IFERR
    Dim strSQL As String
    Dim rs As RecordSet
    
    If Adodc1.RecordSet.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Set rs = New RecordSet
    strSQL = "Select * From G_SGGridMergeCells Where B_ObjectID='" & mvarObjectID & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Do While Not rs.EOF
        SGGrid1.Columns(Trim(rs("B_FieldName"))).MergeCells = sgMergeFree
        
        SGGrid1.Columns(Trim(rs("B_FieldName"))).Style.TextAlignment = sgAlignCenterCenter
        
        
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "本错误发生于：SetSGGridMergeCells" & vbNewLine
    szErr = szErr & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
    
End Sub

Private Sub SetInventoryVoid()
    Dim lID As Long
    lID = SGGrid1.Rows.Current.Cells(SGGrid1.Columns("B_ID").Position).Value
    
    strSQL = "exec dbo.P_SetCPInventoryVoid '" & lID & "'"
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL


    '刷新网格
    PriviewGrid
End Sub



'在2016年12月26日 13:01:39制作的替换掉原来的PrintPreview
'需要在打印预览的时候对特性控件进行修改
Private Sub PrintPreviewEx()
    Dim rpt1 As New ActiveReport1
    Dim cls1 As New clsPrint
    
    
    Dim szReportFile As String
    szReportFile = cls1.DownloadReport(m_ReportObjectID)
    rpt1.Caption = "打印预览"
    rpt1.WindowState = 2
    With rpt1
        .Refresh
        Set .DataControl1.RecordSet = Adodc1.RecordSet.Clone
        .LoadLayout szReportFile
        .Show
    End With
    
    Set rpt1 = Nothing
End Sub

Private Sub PrintPreviewEx01()
    Dim frm1 As New frmModBLRPreviewOri
    frm1.ObjectID = m_ReportObjectID
    Set frm1.RecordSet = Adodc1.RecordSet.Clone
    
    frm1.Show
End Sub

