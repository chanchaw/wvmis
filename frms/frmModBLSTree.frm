VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModBLSTree 
   Caption         =   "Form2"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "frmModBLSTree.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8130
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8130
      _LayoutVersion  =   1
      _ExtentX        =   14340
      _ExtentY        =   10186
      _DataPath       =   ""
      Bands           =   "frmModBLSTree.frx":038A
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3300
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModBLSTree.frx":6234
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModBLSTree.frx":67CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModBLSTree.frx":6D68
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2280
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4515
         Left            =   240
         TabIndex        =   1
         Top             =   1140
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   7964
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList2"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmModBLSTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_SQL As String
Private m_EditFormName As String
Private clsSGridShow1 As clsSGridShow
Private mvarObjectID As String '局部复制
Private sFilter As String
Private iKeyIndex As Integer
Public m_KeyID As Variant

Dim m_TableName As String
Dim m_FieldID As String
Dim m_ParentID As String

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

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "新增"
            AddNewObject
        Case "编辑"
            EditObject m_KeyID
        Case "删除"
            DeleteObject
        Case "刷新"
            LoadObject
        Case "关闭"
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()
    ActiveBar21.ClientAreaControl = TreeView1
    ActiveBar21.RecalcLayout
    GetObjectParameter
    Me.Left = 0
    Me.Top = 0
    
    AnimateForm Me
End Sub

'新增对象
Private Sub AddNewObject()
    On Error Resume Next
    '刷新网格
    Dim o As Object
    
    '判断是否有新增的权限
    
    If Gm.PI.JudgeNew(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    
    Set o = GetFormNew(m_EditFormName)
    With o
        .AddNewObject ObjectID
        .Show vbModal
    End With
    LoadObject
    
    If Adodc1.RecordSet.RecordCount > 0 Then
        Adodc1.RecordSet.MoveLast
    End If
End Sub

'编辑对象
Private Sub EditObject(ByVal m_KeyID As Variant)
    On Error GoTo IFERR
    Dim o As Object
    Dim Irow As Long
    Dim sKey As Variant
    
        '判断是否有新增的权限
    If Gm.PI.JudgeUpdate(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    

    sKey = Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
    
    If sKey = "0" Then
        Exit Sub
    End If

    Set o = GetFormNew(m_EditFormName)
    With o

        .m_KeyID = sKey
        .EditObject ObjectID
        .Show vbModal
    End With
    LoadObject
    
    NavigatorNode sKey
    Exit Sub
IFERR:

End Sub

Private Sub DeleteObject()
    On Error GoTo IFERR
    Dim strSQL As String
    Dim sKey As Variant
    Dim rs As New RecordSet
    
    '判断是否有新增的权限
    If Gm.PI.JudgeDelete(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    
    sKey = Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
    '判定是否可删除
    strSQL = "Select Top 1 * From " & m_TableName & " Where " & m_ParentID & "='" & Trim(sKey) & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "以下还有子分类,不能进行删除", vbExclamation, "删除"
    Else
        If MsgBox("是否要删除?", vbExclamation + vbOKCancel + vbDefaultButton2, "删除") = vbOK Then
            strSQL = "Delete From " & m_TableName & " Where " & m_FieldID & "='" & Trim(sKey) & "'"
            Gm.cnnTool.cnn.Execute strSQL
            
            LoadObject
        End If
    End If
    rs.Close
    Set rs = Nothing
    
    Exit Sub
IFERR:
    Dim szTip As String
    szTip = "存在对应业务数据，不可删除！"
    MsgBox szTip, vbOKOnly + vbInformation, "提示"
    
End Sub

'取得记录
Public Sub LoadObject()

    
    With Adodc1
        .ConnectionString = Gm.cnnTool.cnnStr
        .CommandType = adCmdText
        .RecordSource = m_SQL
        .Refresh
        
    End With
    FillTreeView
    
    
    Me.WindowState = 2
End Sub

Public Function BuillTreeView(ByVal nAreaID As String)
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim Nodx As Node
    
    Set rs = New RecordSet
    
    strSQL = m_SQL & " And B_Parent='" & nAreaID & "'"
    
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        Exit Function
    End If

    Do While Not rs.EOF
        Set Nodx = TreeView1.Nodes.Add("F" & Trim(nAreaID), tvwChild, "F" & Trim(rs(m_FieldID)), rs(m_FieldID), 1, 2)
        Nodx.Expanded = True
    
        Call BuillTreeView(rs.Fields(m_FieldID).Value)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub FillTreeView()
    '取得记录集
    Dim Nodx As Node
    TreeView1.Nodes.Clear
    Set Nodx = TreeView1.Nodes.Add(, tvwFirst, "F0", "全部" & Me.Caption, 3, 3)
    Nodx.Expanded = True
    BuillTreeView "0"
    
    TreeView1.Font.Size = 12
End Sub

Private Sub NavigatorNode(ByVal m_Key As String)
    On Error Resume Next
    Dim o As Node
    For Each o In TreeView1.Nodes
        If Mid(o.Key, 2, Len(o.Key) - 1) = m_Key Then
            o.Selected = True
            Exit Sub
        End If
    Next
End Sub

'取得参数
Private Sub GetField()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = m_SQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    
    rs.AddNew
    
    m_FieldID = rs.Fields(0).name
    m_ParentID = rs.Fields(1).name

    m_TableName = rs.Fields(0).Properties(1).Value
    
    rs.CancelUpdate
    rs.Close
    Set rs = Nothing
End Sub

'取得参数
Private Sub GetObjectParameter()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLS Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly

    
    
    m_SQL = rs("B_SQL")
    m_EditFormName = rs("B_EditFormName")
    
    Me.Width = rs("B_Width")
    Me.Height = rs("B_Height")
    Me.Caption = rs("B_BillName")
    
    rs.Close
    Set rs = Nothing
    
    GetField
End Sub

Private Sub TreeView1_DblClick()
    On Error Resume Next
    Dim sKey As Variant

    sKey = Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
    EditObject sKey
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar21.Bands("Band3").PopupMenu
    End If
    Exit Sub
End Sub
