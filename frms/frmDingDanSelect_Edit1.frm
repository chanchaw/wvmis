VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmDingDanSelect_Edit1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置明细"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
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
   ScaleHeight     =   4200
   ScaleWidth      =   5970
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5970
      _LayoutVersion  =   1
      _ExtentX        =   10530
      _ExtentY        =   7408
      _DataPath       =   ""
      Bands           =   "frmDingDanSelect_Edit1.frx":0000
      Begin XtremeSuiteControls.FlatEdit FlatEdit3 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2880
         Width           =   2535
         _Version        =   1048578
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1860
         Width           =   2535
         _Version        =   1048578
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   900
         Width           =   2535
         _Version        =   1048578
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   2940
         Width           =   975
         _Version        =   1048578
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "注意事项:"
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   975
         _Version        =   1048578
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "毛丝规格:"
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   960
         Width           =   855
         _Version        =   1048578
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "投份:"
      End
   End
End
Attribute VB_Name = "frmDingDanSelect_Edit1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public item As String
Public itemidb As String

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "保存并复制本订单"
            Saveandcopy
        Case "保存"
            save
        Case "退出"
            Unload Me
    
    End Select
End Sub

Private Sub save()
    Dim sql As String
    sql = "update G_BillDetailwhite set B_CastQty='" & FlatEdit1.Text & "',B_Maospecification='" & FlatEdit2.Text & "',B_Attention='" & FlatEdit3.Text & "' where B_Itemid='" & item & "'"
    Gm.cnnTool.cnn.Execute sql
    Me.Hide
End Sub

Private Sub Saveandcopy()
    Dim sql As String
    sql = "update G_BillDetailwhite set B_CastQty='" & FlatEdit1.Text & "',B_Maospecification='" & FlatEdit2.Text & "' where B_itemidb='" & itemidb & "' "
    Gm.cnnTool.cnn.Execute sql
    Me.Hide
End Sub

