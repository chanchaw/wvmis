VERSION 5.00
Begin VB.Form frmColorJRKprintDetail_Write 
   Caption         =   "设置空加值"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4035
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确 定"
         Height          =   600
         Left            =   2520
         TabIndex        =   4
         Top             =   2880
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         Caption         =   "取消退出"
         Height          =   600
         Left            =   480
         TabIndex        =   3
         Top             =   2880
         Width           =   1110
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "空加码数："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "空加米数："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "空加公斤："
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmColorJRKprintDetail_Write"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_Judeg As Long
Public m_TXT1 As String
Public m_TXT2 As String
Public m_TXT3 As String
'退出
Private Sub Command1_Click()
  Unload Me
  
End Sub
'确定
Private Sub Command2_Click()
m_Judeg = 1
m_TXT1 = Text1.Text
m_TXT2 = Text2.Text
m_TXT3 = Text3.Text
 Unload Me
End Sub

Private Sub Form_Load()
m_Judeg = 0
End Sub
