VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmSchedule_Edit2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置完成"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6975
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   660
      Width           =   615
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   840
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
      _Version        =   1048578
      _ExtentX        =   3836
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "保存"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   840
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
      _Version        =   1048578
      _ExtentX        =   3836
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "取消退出"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   1740
      TabIndex        =   2
      Top             =   660
      Width           =   1455
      _Version        =   1048578
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "是否完成:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSchedule_Edit2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bool As Boolean


Private Sub PushButton2_Click()
    
    bool = True
    Me.Hide
    
End Sub

Private Sub PushButton3_Click()
    Unload Me
End Sub
