VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmLoginSOB 
   BorderStyle     =   0  'None
   Caption         =   "登录"
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoginSOB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoginSOB.frx":058A
   ScaleHeight     =   422
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
      _Version        =   1048578
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "登录"
      Appearance      =   2
      ImageGap        =   8
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3510
      Width           =   2775
      _Version        =   1048578
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit3 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4080
      Width           =   2775
      _Version        =   1048578
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
      _Version        =   1048578
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "退出"
      Appearance      =   2
      ImageGap        =   8
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2955
      Width           =   2775
      _Version        =   1048578
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox ComboBox1 
      Height          =   360
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
      _Version        =   1048578
      _ExtentX        =   4895
      _ExtentY        =   635
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   4140
      Width           =   855
      _Version        =   1048578
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "密  码："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   3570
      Width           =   855
      _Version        =   1048578
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "用户名："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   7
      Top             =   3015
      Width           =   975
      _Version        =   1048578
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "工  号："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2460
      Width           =   855
      _Version        =   1048578
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "账  套:"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmLoginSOB"
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

'**************************************************
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
  
'Private Const GWL_STYLE = (-16)
  
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const SubSystem As String = "织造企业MIS系统"
  
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
           Dim ReturnVal As Long
           X = ReleaseCapture()
           ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub
'**************************************************



Private Sub FlatEdit2_Change()
    Dim rs As RecordSet
    Dim sql As String
    Set rs = New RecordSet
    sql = "select B_UserDes from G_SystemUser where B_username='" & Trim(FlatEdit2.Text) & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        FlatEdit1.Text = rs!B_UserDes
    Else
        FlatEdit1.Text = ""
    End If
    
End Sub

Private Sub FlatEdit2_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit3.SetFocus
    End Select
End Sub
Private Sub FlatEdit3_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            denglu
    End Select
End Sub


Private Sub Form_Load()
    
    
'    DisableX Me
'
'    ActiveBar21.Bands("Band1").Tools("Tool0").hwnd = ComBox1.hwnd
'    ActiveBar21.Bands("Band2").Tools("Tool1").hwnd = Combo2.hwnd
'    ActiveBar21.Bands("Band3").Tools("Tool2").hwnd = UCTextBox1.hwnd
'    ActiveBar21.RecalcLayout
'    Label2.Caption = "请输入密码以登录系统!"
    
    
    '初始化登录窗体上为所有子系统
    'InitCombo
    
'    ShowOneSubSystem "织造企业MIS系统"

    
    '记忆最后一次登录的系统的用户名
    Dim m_UserName As String
    Dim m_SubSystem As String
    m_SubSystem = GetSetting(App.Title, "Settings", "SubSystem")
    m_UserName = GetSetting(App.Title, "Settings", "UserName")
'
'    ComBox1.Text = m_SubSystem
'    Combo2.Text = m_UserName
'
    
    'AnimateForm Me
    zhangt
End Sub

Private Sub PushButton1_Click(Index As Integer)
    Select Case Index
    
        Case 0
            denglu
        Case 1
             OK = False
            Me.Hide
    
    End Select
End Sub

Private Sub denglu()
    If Len(Trim$(FlatEdit2.Text)) <= 0 Then
        MsgBox "工号不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
     If Len(Trim(FlatEdit1.Text)) <= 0 Then
        MsgBox "用户名不存在", vbInformation, "提示"
        Exit Sub
    End If
    
'     If Len(Trim$(FlatEdit3.Text)) <= 0 Then
'        MsgBox "密码不能为空", vbInformation, "提示"
'        Exit Sub
'    End If
    Dim rs As RecordSet
    Dim sql As String
    Set rs = New RecordSet
    sql = "select B_PassWord from G_SystemUser where B_username='" & Trim(FlatEdit2.Text) & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If Trim(FlatEdit3.Text) <> rs!B_PassWord Then
        MsgBox "密码错误", vbInformation, "提示"
        Exit Sub
    Else
        stse
        
        OK = True
        Me.Hide
    End If
  
End Sub

Private Sub stse()
        Gm.SysID.SubSystem = SubSystem
        Gm.SysID.SystemUser = FlatEdit2.Text
        Gm.SysID.SystemUserName = FlatEdit1.Text
'        Gm.OnlyDataBreak = IIf(IsNull(rs!B_OnlyDataBreak), 0, rs!B_OnlyDataBreak)
        
        'clsSParameter1.SetParameterString "PictureName", IIf(IsNull(rs("B_PictureName")), "", rs("B_PictureName"))
'        CheckUser = True
        
        
        '保存登录的子系统名称和用户名
        SaveSetting App.Title, "Settings", "SubSystem", ComboBox1.Text
        SaveSetting App.Title, "Settings", "UserName", FlatEdit1.Text
        
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
End Sub

Private Sub zhangt()
    Dim rs As RecordSet
    Set rs = New RecordSet
    Dim sql As String
    sql = "select B_AccountID from G_SetOfBooks Order by B_Order"
    rs.Open sql, Gm.cnnToolSOB.cnn, adOpenKeyset, adLockPessimistic

    
    '绑定账套到下拉控件上
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    ComboBox1.Clear
    rs.MoveFirst
    Do While Not rs.EOF
        ComboBox1.AddItem rs!B_AccountID
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    '默认一个账套
    ComboBox1.ListIndex = 0
End Sub
Private Sub FlatEdit2_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
