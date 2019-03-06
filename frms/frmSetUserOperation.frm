VERSION 5.00
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{332B766E-0D0F-451B-B35F-358EC95AC208}#1.0#0"; "UCCommonCtls.ocx"
Begin VB.Form frmSetUserOperation 
   BackColor       =   &H00CEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户设置"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmSetUserOperation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4305
   StartUpPosition =   2  '屏幕中心
   Begin TA_UCCommonCtls.UCCheckBox UCCheckBox1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1380
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "管理员"
      Text            =   ""
      BackColor       =   13557726
      EdgeHeight      =   180
   End
   Begin TA_UCButton.UCButton UCButton1 
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   2280
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      Caption         =   "确定　"
      Icon            =   "frmSetUserOperation.frx":058A
      IconMask        =   "frmSetUserOperation.frx":0B24
      CaptionAlignment=   1
   End
   Begin TA_UCCommonCtls.UCTextBox UCTextBox1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      TextHeight      =   255
      TextHeight      =   180
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "用户名"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "用户名"
      BackColor       =   -2147483643
      TextHeight      =   255
      CaptionBcckColor=   13557726
      BorderColor     =   16777215
   End
   Begin TA_UCButton.UCButton UCButton2 
      Height          =   435
      Left            =   2220
      TabIndex        =   2
      Top             =   2280
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      Caption         =   "取消退出"
      Icon            =   "frmSetUserOperation.frx":10BE
      IconMask        =   "frmSetUserOperation.frx":1354
      CaptionAlignment=   1
   End
   Begin TA_UCCommonCtls.UCTextBox UCTextBox2 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   420
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      TextHeight      =   255
      TextHeight      =   180
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "用户编号"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "用户编号"
      BackColor       =   -2147483643
      TextHeight      =   255
      CaptionBcckColor=   13557726
      BorderColor     =   16777215
   End
End
Attribute VB_Name = "frmSetUserOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Public m_UserName As String

'初始化角色树形结构的控件
'Private Sub InitRole()
'    With UCTreeViewList1
'        .ConnectionString = Gm.cnnTool.cnnStr
'        .sql = "SELECT B_RoleID, B_Parent FROM G_Role WHERE 1=1"
'        .Refresh
'    End With
'End Sub

Private Sub Form_Load()
'    Dim strSQL As String
'    strSQL = " Select '全部部门' as B_Department"
'    strSQL = strSQL & " Union All"
'    strSQL = strSQL & " Select B_Department From G_Department"
'
'    With UCListBox1
'        .ConnectionString = cnn.cnnStr
'        .SQL = strSQL
'        .Refresh
'
'    End With
    
'    strSQL = " Select '全部业务分组' as B_Team"
'    strSQL = strSQL & " Union All"
'    strSQL = strSQL & " Select B_Team From G_Team"
'
'    With UCListBox2
'        .ConnectionString = cnn.cnnStr
'        .SQL = strSQL
'        .Refresh
'
'    End With
'
'    strSQL = " Select '全部仓库' as B_StorageID"
'    strSQL = strSQL & " Union All"
'    strSQL = strSQL & " Select B_StorageID From G_Storage"
'
'    With UCListBox3
'        .ConnectionString = cnn.cnnStr
'        .SQL = strSQL
'        .Refresh
'
'    End With
    LoadObject
    AnimateForm Me
End Sub

Private Sub LoadObject()
    '初始化角色
'    InitRole

    Dim strSQL As String
    Dim rs As New RecordSet
    Dim szRole As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_SystemUser Where B_UserName='" & m_UserName & "'"
    
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    
    If Not rs.EOF Then
        UCTextBox2.Text = rs("B_UserName")
        UCTextBox1.Text = rs("B_UserDes")
        UCCheckBox1.Value = rs("B_SuperAdmin")
'        szRole = IIf(IsNull(rs!B_Role), "", rs!B_Role)
'        If Len(szRole) > 0 Then
'            UCTreeViewList1.Text = szRole
'        End If
        
        'UCListBox1.Text = rs("B_Department")
        'UCListBox2.Text = rs("B_Team")
        'UCListBox3.Text = rs("B_StorageID")
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub UCButton1_Click()
    If AddNewObject = True Then
        OK = True
        Me.Hide
    End If
End Sub

Private Sub UCButton2_Click()
    OK = False
    Me.Hide
End Sub

Private Function AddNewObject() As Boolean
    On Error GoTo IFERR
    'sUserName = InputBox("请输入要增加的用户名", "用户名称", "", (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2)
    Dim strSQL As String
    Dim sql As String
    Dim rs As New RecordSet
    If Len(m_UserName) < 1 Then
        sql = "select * from G_SystemUser where B_username='" & UCTextBox2.Text & "' "
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            MsgBox "已经存在此编号", vbInformation, "提示"
            Exit Function
        End If
        
        
        strSQL = "Insert Into G_SystemUser (B_username,B_Userdes,B_Password,B_SuperAdmin)"
        strSQL = strSQL & " Values"
        strSQL = strSQL & " ('" & UCTextBox2.Text & "','" & UCTextBox1.Text & "','','" & UCCheckBox1.Value & "')"
        Gm.cnnTool.cnn.Execute strSQL
        
        '新建信箱
        strSQL = "Insert Into G_MsgBox (B_MsgBoxName,B_MsgBoxType,B_UserName)"
        strSQL = strSQL & " Values"
        strSQL = strSQL & " ('收件箱',1,'" & UCTextBox1.Text & "')"
        
        Gm.cnnTool.cnn.Execute strSQL
        
        strSQL = "Insert Into G_MsgBox (B_MsgBoxName,B_MsgBoxType,B_UserName)"
        strSQL = strSQL & " Values"
        strSQL = strSQL & " ('发件箱',2,'" & UCTextBox1.Text & "')"
        
        Gm.cnnTool.cnn.Execute strSQL
    
        strSQL = "Insert Into G_MsgBox (B_MsgBoxName,B_MsgBoxType,B_UserName)"
        strSQL = strSQL & " Values"
        strSQL = strSQL & " ('废件箱',3,'" & UCTextBox1.Text & "')"
        
        Gm.cnnTool.cnn.Execute strSQL
    Else
        'strSQL = "Update G_SystemUser "
        'strSQL = strSQL & " Set B_Department='" & UCListBox1.Text & "'" ',"
        'strSQL = strSQL & " B_Team='" & UCListBox2.Text & "',"
        'strSQL = strSQL & " B_StorageID='" & UCListBox3.Text & "'"
        'strSQL = strSQL & " Where B_UserName='" & UCTextBox1.Text & " '"
        'Gm.cnnTool.cnn.Execute strSQL
        If m_UserName <> UCTextBox2.Text Then
            Dim sql1 As String
            Dim rs1 As New RecordSet
            sql1 = "select * from G_SystemUser where B_username='" & UCTextBox2.Text & "'"
            rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

            If rs1.RecordCount > 0 Then
                MsgBox "已经存在这个编号", vbInformation, "提示"
                Exit Function
            End If
        End If
        strSQL = "Update G_SystemUser "
        strSQL = strSQL & " Set B_SuperAdmin='" & UCCheckBox1.Value & "',B_UserName='" & UCTextBox2.Text & "'"
        strSQL = strSQL & ",B_userdes='" & UCTextBox1.Text & "' Where B_userName='" & UCTextBox2.Text & " '"
        Debug.Print strSQL
        Gm.cnnTool.cnn.Execute strSQL
        
        
    End If
    AddNewObject = True
    Exit Function
IFERR:
    AddNewObject = False
    MsgBox Err.Description
End Function
