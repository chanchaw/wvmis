VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{B45CF12E-4E4F-487D-8096-DB3BFE63F435}#1.0#0"; "sg20ou.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBLSearch 
   BackColor       =   &H00CEDFDE&
   Caption         =   "业务查询"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "frmBLSearch.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   8130
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6540
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8130
      _cx             =   14340
      _cy             =   11536
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
      GridRows        =   2
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmBLSearch.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   5145
         ScaleHeight     =   585
         ScaleWidth      =   2985
         TabIndex        =   4
         Top             =   5955
         Width           =   2985
         Begin TA_UCButton.UCButton UCButton2 
            Height          =   375
            Left            =   1620
            TabIndex        =   6
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   " 取消 "
            Icon            =   "frmBLSearch.frx":03D7
            IconMask        =   "frmBLSearch.frx":066D
            CaptionAlignment=   1
         End
         Begin TA_UCButton.UCButton UCButton1 
            Height          =   375
            Left            =   180
            TabIndex        =   5
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "打开 "
            Icon            =   "frmBLSearch.frx":0903
            IconMask        =   "frmBLSearch.frx":0C9D
            CaptionAlignment=   1
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   0
         ScaleHeight     =   585
         ScaleWidth      =   5145
         TabIndex        =   3
         Top             =   5955
         Width           =   5145
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
         Height          =   5955
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8130
         _LayoutVersion  =   1
         _ExtentX        =   14340
         _ExtentY        =   10504
         _DataPath       =   ""
         Bands           =   "frmBLSearch.frx":1037
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   2820
            TabIndex        =   8
            Top             =   1560
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Format          =   264896513
            CurrentDate     =   38699
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1260
            TabIndex        =   7
            Top             =   1560
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Format          =   264896513
            CurrentDate     =   38699
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   5040
            Top             =   2700
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   661
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
         Begin DDSharpGridOLEDB2U.SGGrid SGGrid1 
            Bindings        =   "frmBLSearch.frx":31F1
            Height          =   5355
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   7875
            _cx             =   13891
            _cy             =   9446
            DataMember      =   ""
            DataMode        =   1
            AutoFields      =   -1  'True
            Enabled         =   -1  'True
            GridBorderStyle =   1
            ScrollBars      =   3
            FlatScrollBars  =   1
            ScrollBarTrack  =   0   'False
            DataRowCount    =   2
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
            TextAlignment   =   0
            WordWrap        =   0   'False
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
            SkipReadOnly    =   -1  'True
            DefaultColWidth =   1200
            DefaultRowHeight=   225
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
            MultiSelect     =   1
            AllowAddNew     =   0   'False
            AllowDelete     =   0   'False
            AllowEdit       =   0   'False
            ScrollBarTips   =   0
            CellTips        =   0
            CellTipsDelay   =   1000
            SpecialMode     =   1
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
            Format          =   "frmBLSearch.frx":3206
            Caption         =   "frmBLSearch.frx":3238
            ScrollTipColumn =   "frmBLSearch.frx":325C
            GroupByBoxText  =   "frmBLSearch.frx":3280
            StylesCollection=   "frmBLSearch.frx":3306
            ColumnsCollection=   "frmBLSearch.frx":6EAC
            ValueItems      =   "frmBLSearch.frx":8B7E
         End
      End
   End
End
Attribute VB_Name = "frmBLSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frm1 As Object
Public m_ObjectID As String

Dim m_SearchSQL As String
Dim clsSGridShow1 As New clsSGridShow

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "刷新"
            PriviewGrid
        Case "收缩"
            SGGrid1.CollapseAll
        Case "展开"
            SGGrid1.ExpandAll
        Case "过滤"
            FilterForm
            
        Case "导出"
            SaveToExcelU Adodc1.RecordSet
        Case "关闭"
            Unload Me
    End Select
End Sub

'另存到Excel中
Public Sub SaveToExcelU(ByRef rs As RecordSet)
    On Error GoTo IFERR
    Dim Irow, Icol As Integer
    Dim Irowcount, Icolcount As Integer
    Dim Fieldlen() '存字段长度值
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    Irowcount = rs.RecordCount
    Icolcount = rs.Fields.Count
    ReDim Fieldlen(Icolcount)
    xlApp.Visible = True '显示表格
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        xlSheet.Cells(1, i + 1).Value = GetCnName(rs.Fields(i).name)
    Next


    Irow = 1
    Do While Not rs.EOF
        For Icol = 0 To Icolcount - 1
            xlSheet.Cells(Irow + 1, Icol + 1).NumberFormatLocal = "@"
            xlSheet.Cells(Irow + 1, Icol + 1).Value = rs(Icol)
        Next
        
        Irow = Irow + 1
        rs.MoveNext
    Loop

    xlApp.Visible = True '显示表格
    'xlBook.Save '保存"
    Set xlApp = Nothing '交还控制给Excel
    Exit Sub
IFERR:
    MsgBox "Excel导出时不正确!", vbExclamation, "Excel"
    Exit Sub
End Sub

Private Function GetCnName(ByVal m_FieldName As String) As String
    Dim strSQL As String
    Dim rs As New RecordSet
    Set rs = New RecordSet
    strSQL = "Select B_CnName From G_FieldUser Where B_FieldName='" & Trim(m_FieldName) & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        GetCnName = rs(0)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub ActiveBar21_ToolComboClose(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    PriviewGrid
End Sub


Private Sub Form_Load()
    With Me
        .Width = 9500
        .Height = 7000
    End With
    DTPicker2.Value = Now
    DTPicker1.Value = Now - 30
    
    ActiveBar21.ClientAreaControl = SGGrid1
    ActiveBar21.Bands("菜单").Tools("Tool0").hwnd = DTPicker1.hwnd
    ActiveBar21.Bands("菜单").Tools("Tool1").hwnd = DTPicker2.hwnd
    
    ActiveBar21.Bands("菜单").Tools("状态").CBAddItem "全部状态"
    ActiveBar21.Bands("菜单").Tools("状态").CBAddItem "已登帐"
    ActiveBar21.Bands("菜单").Tools("状态").CBAddItem "草稿"
    ActiveBar21.Bands("菜单").Tools("状态").CBListIndex = 0
    ActiveBar21.RecalcLayout
    
    GetObjectParameter
    PriviewGrid
    AnimateForm Me
    
    
End Sub

Private Sub PriviewGrid()
    On Error GoTo IFERR
    Dim strSQL As String
    strSQL = m_SearchSQL & " And B_Date Between '" & Trim(Format(DTPicker1.Value, "YYYY-MM-DD")) & "' And '" & Trim(Format(DTPicker2.Value, "YYYY-MM-DD")) & "'"
    
    Select Case ActiveBar21.Bands("菜单").Tools("状态").Text
        Case "已登帐"
            strSQL = strSQL & " And B_State='已登帐'"
        Case "草稿"
            strSQL = strSQL & " And B_State='草稿'"
    End Select
    
    With Adodc1
        '.ConnectionString = cnn.cnnstr
        .ConnectionString = Gm.cnnTool.cnnStr
        .CommandType = adCmdText
        .RecordSource = strSQL
        .Refresh
        
    End With

    ShowSGGrid

    Exit Sub
IFERR:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub ShowSGGrid()
    Set clsSGridShow1 = New clsSGridShow
    With clsSGridShow1
        .ObjectID = m_ObjectID
        .InitClass SGGrid1, 2
        .FillGrid Adodc1.RecordSet
'        SGGrid1.DataMode = sgBound
'        Set SGGrid1.DataSource = Adodc1
'        SGGrid1.ReBind
        .ShowGridFormat
        .GroupGrid "B_CodeID", ""
        
    End With
End Sub

Private Sub FilterForm()
    Dim frm1 As New frmUCTSearch
    With frm1
        Set .rs = Adodc1.RecordSet
        .ObjectID = m_ObjectID
        .FieldType = 4
        .Show vbModal
    End With
    If frm1.OK = True Then
        Adodc1.RecordSet.Filter = ""
        Adodc1.RecordSet.Filter = frm1.strResult
        ShowSGGrid
    End If
    Unload frm1
    Set frm1 = Nothing
End Sub

'取得参数
Private Sub GetObjectParameter()
    On Error Resume Next
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select B_SearchSQL From G_BL Where B_ObjectID='" & m_ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly

    m_SearchSQL = rs("B_SearchSQL")
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub SGGrid1_DblClick()
    OpenFrame
End Sub

Private Sub UCButton1_Click()
    OpenFrame
End Sub

Private Sub OpenFrame()
    If Val(SGGrid1.CurrentCell.row.Cells(1).Value) > 0 Then
        frm1.Refresh
        frm1.EditObject Val(SGGrid1.CurrentCell.row.Cells(1).Value)
    End If
End Sub

Private Sub UCButton2_Click()
    Unload Me
End Sub
