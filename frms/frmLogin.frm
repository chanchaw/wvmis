VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{332B766E-0D0F-451B-B35F-358EC95AC208}#1.0#0"; "UCCommonCtls.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00CEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "系统登录"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   DrawMode        =   15  'Merge Pen Not
   DrawStyle       =   2  'Dot
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5115
   StartUpPosition =   2  '屏幕中心
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   3210
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5115
      _cx             =   9022
      _cy             =   5662
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
      BackColor       =   13557726
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
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmLogin.frx":058A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
         Height          =   1695
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   5115
         _LayoutVersion  =   1
         _ExtentX        =   9022
         _ExtentY        =   2990
         _DataPath       =   ""
         Bands           =   "frmLogin.frx":05D3
         Begin TA_UCCommonCtls.UCTextBox UCTextBox1 
            Height          =   435
            Left            =   0
            TabIndex        =   0
            Top             =   1200
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   767
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
            Caption         =   "密码:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "密码:"
            BackColor       =   -2147483643
            TextHeight      =   255
            CaptionBcckColor=   13557726
            PasswordChar    =   "*"
            BorderColor     =   16777215
         End
         Begin TA_UCCommonCtls.UCComBox ComBox1 
            Height          =   435
            Left            =   0
            TabIndex        =   9
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   767
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "子系统:"
            Picture1.Backcolor=   13557726
         End
         Begin TA_UCCommonCtls.UCComBox Combo2 
            Height          =   435
            Left            =   0
            TabIndex        =   8
            Top             =   720
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   767
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "用户名:"
            Picture1.Backcolor=   13557726
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00F7FFFF&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   0
         ScaleHeight     =   720
         ScaleWidth      =   5115
         TabIndex        =   5
         Top             =   0
         Width           =   5115
         Begin VB.Label Label1 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "欢迎使用"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   360
            TabIndex        =   6
            Top             =   300
            Width           =   4455
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   5115
         TabIndex        =   2
         Top             =   2415
         Width           =   5115
         Begin TA_UCButton.UCButton cmdCancel 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "取消  "
            Icon            =   "frmLogin.frx":0F63
            IconMask        =   "frmLogin.frx":11F9
            CaptionAlignment=   1
         End
         Begin TA_UCButton.UCButton cmdOK 
            Height          =   375
            Left            =   1980
            TabIndex        =   4
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "确定  "
            Icon            =   "frmLogin.frx":148F
            IconMask        =   "frmLogin.frx":1829
            CaptionAlignment=   1
         End
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H8

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_DISABLED = &H2&

Dim clsEcode1 As New clsEcode

Public Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub

Public Sub SetFormTopmost(theForm As Form)

SetWindowPos theForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE

End Sub

Private Sub DisableX(ByRef frm As Form)
    Dim hMenu As Long, nCount As Long
    hMenu = GetSystemMenu(frm.hwnd, 0)
    nCount = GetMenuItemCount(hMenu)
    Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
    DrawMenuBar frm.hwnd
End Sub


Private Sub cmdOK_Click()
    Gm.SysID.SubSystem = ComBox1.Text
    If CheckUser = True Then
        OK = True
        Me.Hide
    End If
End Sub

Private Sub ComBox1_Change()
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim sList As String
    
    Set rs = New RecordSet
    strSQL = "exec dbo.usp_GetSubSystemUser '" & Trim(ComBox1.Text) & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    'sList = "管理员,"
    Do While Not rs.EOF
        sList = sList & rs(0) & ","
        rs.MoveNext
    Loop
    
    If Len(sList) > 0 Then
        sList = Mid(sList, 1, Len(sList) - 1)
    Else
        sList = ""
    End If
    
    With Combo2
        .DefaultValue = sList
        .Refresh
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            cmdOK_Click
    End Select
End Sub

Private Sub Form_Load()
    
    
    DisableX Me
    
    ActiveBar21.Bands("Band1").Tools("Tool0").hwnd = ComBox1.hwnd
    ActiveBar21.Bands("Band2").Tools("Tool1").hwnd = Combo2.hwnd
    ActiveBar21.Bands("Band3").Tools("Tool2").hwnd = UCTextBox1.hwnd
    ActiveBar21.RecalcLayout
    Label1.Caption = "请输入密码以登录系统!"
    
    
    '初始化登录窗体上为所有子系统
    'InitCombo
    
    ShowOneSubSystem "织造企业MIS系统"

    
    '记忆最后一次登录的系统的用户名
    Dim m_UserName As String
    Dim m_SubSystem As String
    m_SubSystem = GetSetting(App.Title, "Settings", "SubSystem")
    m_UserName = GetSetting(App.Title, "Settings", "UserName")

    ComBox1.Text = m_SubSystem
    Combo2.Text = m_UserName
    
    
    'AnimateForm Me
    
End Sub


Private Sub InitCombo()
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim sList As String
    
    Set rs = New RecordSet
    
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo

    
    sList = ""
    
    strSQL = "Select B_SubSystem From G_SubSystem Order By B_MenuObjectID"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    
    Do While Not rs.EOF
        sList = sList & rs(0) & ","
        rs.MoveNext
    Loop
    sList = Mid(sList, 1, Len(sList) - 1)
    With ComBox1
        .DefaultValue = sList
        .Refresh
    End With
    rs.Close


    Dim m_SubSystem As String
    m_SubSystem = GetSetting(App.Title, "Settings", "SubSystem")
    If Len(m_SubSystem) > 2 Then
        'ComBox1.Text = m_SubSystem
        
'        ComBox1.Text = "计划科"
'        Combo2.Text = "计划单管理员"
        'ComBox1.Text = "白玉兰印染企业管理"
        'Combo2.Text = "管理员"
    End If

    Set rs = Nothing
End Sub

'仅仅显示一个子系统
Private Function ShowOneSubSystem(ByVal m_SubSystem As String) As Boolean
    Dim strSQL As String
    Dim rs As RecordSet
    Dim sList As String
    
    
    '初始化为TRUE
    ShowOneSubSystem = True
    
    
    Set rs = New RecordSet
    
    'strSQL = "Select * From G_SUserSub Where B_SubSystem='" & m_SubSystem & "' Order By B_ID"
    
    strSQL = "exec dbo.usp_GetSubSystemUser '" & m_SubSystem & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    If rs.RecordCount <= 0 Then
        MsgBox "当前子系统不存在！", vbOKOnly + vbInformation, "提示"
        rs.Close
        Set rs = Nothing
        ShowOneSubSystem = False
        Exit Function
    End If
    
    
    '初始化子系统下拉框
    '===============================
    sList = rs!B_SubSystem
    With ComBox1
        .DefaultValue = sList
        .Refresh
    End With
    
    
    
    '初始化该子系统下的所有用户名
    '===============================
    rs.MoveFirst
    sList = ""
    Do While Not rs.EOF
        sList = sList & rs!B_UserName & ","
        rs.MoveNext
    Loop
    sList = Mid(sList, 1, Len(sList) - 1)
    
    
    With Combo2
        .DefaultValue = sList
        .Refresh
    End With
    '===============================
    
    'msgbox "初始化2个子系统的控件完毕"
    
    rs.Close
    Set rs = Nothing
End Function

Private Function CheckUser() As Boolean
    On Error GoTo IFERR
    Dim sKey As String
    Dim sPassWord As String
    
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    'cnn.InitializeConnection
    
    strSQL = "Select * From G_SystemUser Where B_UserName='" & Trim(Combo2.Text) & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    sPassWord = clsEcode1.DeCode(IIf(IsNull(rs("B_PassWord")), "", rs("B_PassWord")), "ABCDEFGHIJKL")

    If CVar(UCTextBox1.Text) = CVar(sPassWord) Then
        Gm.SysID.SubSystem = ComBox1.Text
        Gm.SysID.SystemUser = Combo2.Text
        Gm.OnlyDataBreak = IIf(IsNull(rs!B_OnlyDataBreak), 0, rs!B_OnlyDataBreak)
        
        'clsSParameter1.SetParameterString "PictureName", IIf(IsNull(rs("B_PictureName")), "", rs("B_PictureName"))
        CheckUser = True
        
        
        '保存登录的子系统名称和用户名
        SaveSetting App.Title, "Settings", "SubSystem", ComBox1.Text
        SaveSetting App.Title, "Settings", "UserName", Combo2.Text
        
        strSQL = "Delete From G_HostLogin"
        strSQL = strSQL & " Where B_HostName='" & Gm.SysID.ComputerName & "'"
        
        Gm.cnnTool.cnn.Execute strSQL
        
        strSQL = "Insert Into G_HostLogin "
        strSQL = strSQL & " ("
        strSQL = strSQL & " B_HostName,B_UserName"
        strSQL = strSQL & " )"
        strSQL = strSQL & " Values"
        strSQL = strSQL & " ("
        strSQL = strSQL & " '" & Gm.SysID.ComputerName & "',"
        strSQL = strSQL & " '" & Gm.SysID.ComputerUserName & "'"
        strSQL = strSQL & " )"
        
        Gm.cnnTool.cnn.Execute strSQL
        
    Else
        MsgBox "口令不正确!", vbExclamation, "登录"
    End If
    Exit Function
IFERR:
    Exit Function
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

End Sub

