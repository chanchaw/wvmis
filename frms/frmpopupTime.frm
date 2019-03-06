VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmpopupTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择时间"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
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
   ScaleHeight     =   3810
   ScaleWidth      =   4155
   StartUpPosition =   2  '屏幕中心
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   555
      Left            =   1620
      TabIndex        =   0
      Top             =   3000
      Width           =   915
      _Version        =   1048578
      _ExtentX        =   1614
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "保存"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   3030
      Width           =   255
      _Version        =   1048578
      _ExtentX        =   450
      _ExtentY        =   556
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   540
      TabIndex        =   1
      Top             =   360
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   201719809
      CurrentDate     =   43060
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   555
      Left            =   2820
      TabIndex        =   4
      Top             =   3000
      Width           =   915
      _Version        =   1048578
      _ExtentX        =   1614
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "退出"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3060
      Width           =   795
      _Version        =   1048578
      _ExtentX        =   1402
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "应用全部"
   End
End
Attribute VB_Name = "frmpopupTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public time As String
Public a As String
Public bsaved As Boolean

Private Sub Form_Load()
    MonthView1.Value = Now
    bsaved = False
End Sub

Private Sub PushButton1_Click()
    time = MonthView1.Value
    a = CheckBox1.Value
    bsaved = True
    Me.Hide
End Sub

Private Sub PushButton2_Click()
    Unload Me
End Sub
