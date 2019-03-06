VERSION 5.00
Begin VB.Form frmDingDanSelect_Edit3 
   Caption         =   "设置色布打卷样式"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDingDanSelect_Edit3.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3690
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "米数+码数"
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "公斤+码数"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "公斤+米数"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "码数"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "米数"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确  定"
         Height          =   600
         Left            =   2040
         TabIndex        =   3
         Top             =   2400
         Width           =   1350
      End
      Begin VB.CommandButton Command1 
         Caption         =   "取消退出"
         Height          =   600
         Left            =   480
         TabIndex        =   2
         Top             =   2400
         Width           =   1230
      End
      Begin VB.OptionButton Option1 
         Caption         =   "公斤"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDingDanSelect_Edit3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_ID As Long  'G_BillDetailColor的主键
Public m_DaJuanGS As String  '用来保存选择的类型
Public bsaved As Boolean


'取消退出
Private Sub Command1_Click()
   Unload Me
End Sub
'确定
Private Sub Command2_Click()
   Dim a As Long
   For a = 0 To 5
        If Option1(a).Value = True Then
          m_DaJuanGS = Option1(a).Caption
        End If
   Next
   bsaved = True
    Unload Me
End Sub

Private Sub Form_Load()
    load
    bsaved = False
End Sub
'刷新
Private Sub load()
Dim a As Long
Dim sql As String
Dim rs As New RecordSet
sql = "SELECT * FROM  G_BillDetailColor WHERE B_ItemID='" & m_ID & "'"
rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

For a = 0 To 5
    Option1(a).Value = False
    If Option1(a).Caption = rs!B_DaJuanGS Then
        Option1(a).Value = True
    End If
Next

End Sub
