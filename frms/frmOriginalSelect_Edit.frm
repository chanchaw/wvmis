VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOriginalSelect_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7485
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
   ScaleHeight     =   4350
   ScaleWidth      =   7485
   StartUpPosition =   2  '屏幕中心
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   855
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
      _Version        =   1048578
      _ExtentX        =   3836
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "确认"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   2430
      Width           =   1635
      _Version        =   1048578
      _ExtentX        =   2884
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
      _Version        =   1048578
      _ExtentX        =   3836
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "退出"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1065
      Width           =   375
      _Version        =   1048578
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   ".."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1065
      Width           =   1335
      _Version        =   1048578
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   14737632
      Enabled         =   0   'False
      BackColor       =   14737632
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit3 
      Height          =   300
      Left            =   1560
      TabIndex        =   7
      Top             =   367
      Width           =   1635
      _Version        =   1048578
      _ExtentX        =   2884
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit4 
      Height          =   300
      Left            =   5040
      TabIndex        =   9
      Top             =   367
      Width           =   1635
      _Version        =   1048578
      _ExtentX        =   2884
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit5 
      Height          =   300
      Left            =   5040
      TabIndex        =   11
      Top             =   1102
      Width           =   1635
      _Version        =   1048578
      _ExtentX        =   2884
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit6 
      Height          =   660
      Left            =   5040
      TabIndex        =   13
      Top             =   1800
      Width           =   1635
      _Version        =   1048578
      _ExtentX        =   2884
      _ExtentY        =   1164
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      MultiLine       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit7 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
      _Version        =   1048578
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   14737632
      Enabled         =   0   'False
      BackColor       =   14737632
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1860
      Width           =   1215
      _Version        =   1048578
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "收货地址："
      ForeColor       =   0
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   1830
      Width           =   795
      _Version        =   1048578
      _ExtentX        =   1402
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "备       注:"
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1125
      Width           =   795
      _Version        =   1048578
      _ExtentX        =   1402
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "送货数量:"
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   390
      Width           =   1035
      _Version        =   1048578
      _ExtentX        =   1826
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "余未分配数量:"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   390
      Width           =   795
      _Version        =   1048578
      _ExtentX        =   1402
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "申请数量:"
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1125
      Width           =   1215
      _Version        =   1048578
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "收货单位："
      ForeColor       =   0
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2460
      Width           =   795
      _Version        =   1048578
      _ExtentX        =   1402
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "单       价:"
   End
End
Attribute VB_Name = "frmOriginalSelect_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bool As Boolean
Public Originalsuppliers As String

'交货方式
Private Sub delivery()
'    Dim sql As String
'    Dim rs As New RecordSet
'    sql = "Select B_SID From G_Delivery Where 1=1"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
'    If rs.RecordCount > 0 Then
'        Do While Not rs.EOF
'            ComboBox2.AddItem "" & rs!B_sid & ""
'            rs.movenext
'        Loop
'    End If
End Sub
'运费结算方式
Private Sub ClearWay()
'    Dim sql As String
'    Dim rs As New RecordSet
'    sql = "Select B_SID From G_Balance Where 1=1"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
'    If rs.RecordCount > 0 Then
'        Do While Not rs.EOF
'            ComboBox3.AddItem "" & rs!B_sid & ""
'            rs.movenext
'        Loop
'    End If
End Sub

Private Sub Form_Load()
    delivery
    ClearWay
    bool = False
'    DTPicker1.Value = Now
'    DTPicker2.Value = Now
End Sub

Private Sub PushButton1_Click()
    
'    If Val(FlatEdit1.Text) <= 0 Then
'        MsgBox "单价不能为0", vbInformation, "提示"
'        Exit Sub
'    End If
'    If Trim(FlatEdit2.Text) = "" Then
'        MsgBox "收货单位不能为空", vbInformation, "提示"
'        Exit Sub
'    End If
    If Val(FlatEdit5.Text) <= 0 Then
        MsgBox "送货数量不能为空", vbInformation, "提示"
        Exit Sub
    End If
    bool = True
    Me.Hide
End Sub

Private Sub PushButton2_Click()
    Unload Me
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
End Sub
Private Sub FlatEdit5_KeyPress(KeyAscii As Integer)
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

Private Sub PushButton3_Click()
    Dim frm1 As New frmPopupDanWei
    frm1.ContactType = "白坯加工商"
    frm1.Caption = "送货地址"
    frm1.Show vbModal
    Originalsuppliers = frm1.Clientid
    FlatEdit2.Text = frm1.ClientName
    ClientName (Originalsuppliers)
    Unload frm1
End Sub

Private Sub ClientName(ByVal a As String)
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_ContactCompany where B_clientid='" & a & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        FlatEdit7.Text = rs!B_Address
    Else
        FlatEdit7.Text = ""
    End If
End Sub
