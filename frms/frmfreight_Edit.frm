VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmfreight_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "生成运费单"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfreight_Edit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9510
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check2 
      Caption         =   "已收回单"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   1500
      Width           =   1215
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   1080
      Left            =   1920
      TabIndex        =   12
      Top             =   4320
      Width           =   2415
      _Version        =   1048578
      _ExtentX        =   4260
      _ExtentY        =   1905
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
   Begin VB.CheckBox Check1 
      Caption         =   "运费已付"
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   1500
      Width           =   1215
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   2340
      TabIndex        =   0
      Top             =   420
      Width           =   315
      _Version        =   1048578
      _ExtentX        =   556
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   ".."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Top             =   420
      Width           =   1335
      _Version        =   1048578
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BackColor       =   14737632
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   375
      Left            =   4260
      TabIndex        =   2
      Top             =   420
      Width           =   1575
      _Version        =   1048578
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit4 
      Height          =   375
      Left            =   7500
      TabIndex        =   3
      Top             =   420
      Width           =   1575
      _Version        =   1048578
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit7 
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      Top             =   1500
      Width           =   1575
      _Version        =   1048578
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit8 
      Height          =   375
      Left            =   1020
      TabIndex        =   5
      Top             =   1500
      Width           =   1575
      _Version        =   1048578
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   1080
      Left            =   5400
      TabIndex        =   13
      Top             =   4320
      Width           =   2415
      _Version        =   1048578
      _ExtentX        =   4260
      _ExtentY        =   1905
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
   Begin XtremeSuiteControls.FlatEdit FlatEdit3 
      Height          =   375
      Left            =   1500
      TabIndex        =   15
      Top             =   2520
      Width           =   2775
      _Version        =   1048578
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit5 
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   2520
      Width           =   2775
      _Version        =   1048578
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit6 
      Height          =   375
      Left            =   1500
      TabIndex        =   19
      Top             =   3480
      Width           =   7455
      _Version        =   1048578
      _ExtentX        =   13150
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   255
      Left            =   300
      TabIndex        =   20
      Top             =   3540
      Width           =   1095
      _Version        =   1048578
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "银行卡开户行:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   2580
      Width           =   1095
      _Version        =   1048578
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "银行卡号:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   300
      TabIndex        =   16
      Top             =   2580
      Width           =   1095
      _Version        =   1048578
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "运费持卡人:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      Top             =   480
      Width           =   855
      _Version        =   1048578
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "车 牌 号:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   315
      Left            =   300
      TabIndex        =   9
      Top             =   450
      Width           =   615
      _Version        =   1048578
      _ExtentX        =   1085
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "运 方:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   3420
      TabIndex        =   8
      Top             =   480
      Width           =   735
      _Version        =   1048578
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "运   费:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Left            =   3420
      TabIndex        =   7
      Top             =   1560
      Width           =   855
      _Version        =   1048578
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "运方电话:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label12 
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   1560
      Width           =   615
      _Version        =   1048578
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "驾驶员:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmfreight_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bool As Boolean
Public Originalsuppliers As String

Private Sub PushButton2_Click()
    If Val(FlatEdit2.Text) <= 0 Then
        MsgBox "运费不能为0", vbInformation, "提示"
    End If
    bool = True
    Me.Hide
    
End Sub

Private Sub PushButton3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    bool = False
End Sub

Private Sub PushButton1_Click()
     Dim frm1 As New frmPopupDanWei
        frm1.ContactType = "物流运输"
        frm1.Caption = "物流运输"
        frm1.TDBGrid1.Columns("B_ClientID").Caption = "物流运输编号"
        frm1.TDBGrid1.Columns("B_ClientName").Caption = "物流运输名称"
        frm1.Show vbModal
        Originalsuppliers = frm1.Clientid
        FlatEdit1.Text = frm1.ClientName
        Unload frm1
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
Private Sub FlatEdit7_KeyPress(KeyAscii As Integer)
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


