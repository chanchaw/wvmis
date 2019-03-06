VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmDingDanSelect_Edit2 
   Caption         =   "设置打卷"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6360
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
   ScaleHeight     =   4410
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   4410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      _LayoutVersion  =   1
      _ExtentX        =   11218
      _ExtentY        =   7779
      _DataPath       =   ""
      Bands           =   "frmDingDanSelect_Edit2.frx":0000
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   2700
         Width           =   1695
         _Version        =   1048578
         _ExtentX        =   2990
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
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit3 
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
         _Version        =   1048578
         _ExtentX        =   2990
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
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72810497
         CurrentDate     =   43110
      End
      Begin XtremeSuiteControls.ComboBox ComboBox2 
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Top             =   960
         Width           =   1695
         _Version        =   1048578
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox ComboBox1 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   1695
         _Version        =   1048578
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2760
         Width           =   855
         _Version        =   1048578
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "公        斤:"
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   1860
         Width           =   855
         _Version        =   1048578
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "匹数:"
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1860
         Width           =   855
         _Version        =   1048578
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "起        期:"
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   960
         Width           =   855
         _Version        =   1048578
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "正次品:"
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   855
         _Version        =   1048578
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "布的种类:"
      End
   End
End
Attribute VB_Name = "frmDingDanSelect_Edit2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bool As Boolean


Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "保存"
            bool = True
            Me.Hide
        Case "退出"
            Unload Me
    
    End Select
End Sub




Private Sub cob1()
    ComboBox1.Clear
    ComboBox1.AddItem "经编布"
    ComboBox1.AddItem "圆机布"
  
    ComboBox1.Text = "经编"
End Sub
Private Sub cob2()
    ComboBox2.Clear
    ComboBox2.AddItem "正品"
    ComboBox2.AddItem "次品"
  
    ComboBox2.Text = "正品"
End Sub
Private Sub Form_Load()
    DTPicker1.Value = Now
    cob1
    cob2
    bool = False
End Sub

Private Sub FlatEdit1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
     FlatEdit1.Text = Format(FlatEdit1.Text, "0.0")
End Sub
Private Sub FlatEdit3_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
     Dim a As Long
     If Len(FlatEdit3.Text) > 0 Then
     a = FlatEdit3.Text
     FlatEdit3.Text = a
     End If
End Sub

