VERSION 5.00
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{332B766E-0D0F-451B-B35F-358EC95AC208}#1.0#0"; "UCCommonCtls.ocx"
Begin VB.Form frmSetPassword 
   BackColor       =   &H00CEDFDE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置口令"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmSetPassword.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TA_UCCommonCtls.UCTextBox Text3 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   873
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
      Caption         =   "重复新口令:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "重复新口令:"
      BackColor       =   -2147483643
      TextHeight      =   255
      CaptionBcckColor=   13557726
      PasswordChar    =   "*"
      BorderColor     =   16777215
   End
   Begin TA_UCCommonCtls.UCTextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   873
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
      Caption         =   "新口令:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "新口令:"
      BackColor       =   -2147483643
      TextHeight      =   255
      CaptionBcckColor=   13557726
      PasswordChar    =   "*"
      BorderColor     =   16777215
   End
   Begin TA_UCCommonCtls.UCTextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   873
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
      Caption         =   "旧口令:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "旧口令:"
      BackColor       =   -2147483643
      TextHeight      =   255
      CaptionBcckColor=   13557726
      PasswordChar    =   "*"
      BorderColor     =   16777215
   End
   Begin TA_UCButton.UCButton UCButton1 
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   3
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "关闭  "
      Icon            =   "frmSetPassword.frx":058A
      IconMask        =   "frmSetPassword.frx":08DC
      CaptionAlignment=   1
   End
   Begin TA_UCButton.UCButton UCButton1 
      Height          =   375
      Index           =   0
      Left            =   3660
      TabIndex        =   4
      Top             =   180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "修改口令"
      Icon            =   "frmSetPassword.frx":0C2E
      IconMask        =   "frmSetPassword.frx":11C8
      CaptionAlignment=   1
   End
End
Attribute VB_Name = "frmSetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsEcode1 As New clsEcode

Private m_UserName As String

Private Sub CheckPassWord()
    On Error Resume Next
    Dim rs As New RecordSet
    Dim mstrSQL As String
    
    'cnn.InitializeConnection
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    Set rs = New RecordSet
    
    mstrSQL = "Select * From G_SystemUser Where B_UserName=" & "'" & Trim(m_UserName) & "'"
    rs.Open mstrSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    If Not rs.EOF Then
'        If clsEcode1.EnCode(Text1.Text, "ABCDEFGHIJKL") = IIf(IsNull(rs("B_Password")), "", rs("B_Password")) Then
        If Text1.Text = IIf(IsNull(rs("B_Password")), "", rs("B_Password")) Then
            If Trim(Text2.Text) = Trim(Text3.Text) Then
            
                'rs("B_Password") = clsEcode1.EnCode(Text2.Text, "ABCDEFGHIJKL")
                rs("B_Password") = Text2.Text
                rs.Update
                MsgBox "已成功修改口令!", vbInformation, "口令"
                Set rs = Nothing
                Unload Me
            Else
                MsgBox "口令不正确!", vbExclamation, "口令"
            End If
        Else
            MsgBox "旧口令不正确!", , "口令"
        End If
    Else
        MsgBox "用户名没有找到!", vbExclamation, "口令"
    End If
    
    rs.Close
    Set rs = Nothing

End Sub


Private Sub Form_Load()
    'm_UserName = clsSParameter1.GetParameterString("UserName")
    m_UserName = Gm.SysID.SystemUser
    AnimateForm Me
End Sub

Private Sub UCButton1_Click(Index As Integer)
    Select Case Index
        Case 0
            CheckPassWord
        Case 1
            Unload Me
    End Select
End Sub
