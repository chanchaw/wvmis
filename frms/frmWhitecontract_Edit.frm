VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmWhitecontract_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "白坯入库"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11175
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
   ScaleHeight     =   4440
   ScaleWidth      =   11175
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   4440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _LayoutVersion  =   1
      _ExtentX        =   19711
      _ExtentY        =   7832
      _DataPath       =   ""
      Bands           =   "frmWhitecontract_Edit.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7530
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   660
         Width           =   12060
         _cx             =   21273
         _cy             =   13282
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   800
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   3
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   5
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmWhitecontract_Edit.frx":0F3C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   6720
            Left            =   45
            ScaleHeight     =   6720
            ScaleWidth      =   11970
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   11970
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   1320
               TabIndex        =   3
               Top             =   2460
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               Format          =   205651969
               CurrentDate     =   43067
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2880
               TabIndex        =   4
               Top             =   300
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   1320
               TabIndex        =   5
               Top             =   300
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   375
               Left            =   4920
               TabIndex        =   6
               Top             =   240
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   1320
               TabIndex        =   7
               Top             =   1320
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               Height          =   375
               Left            =   4920
               TabIndex        =   8
               Top             =   1320
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               Height          =   375
               Left            =   8340
               TabIndex        =   9
               Top             =   1320
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               Height          =   375
               Left            =   4920
               TabIndex        =   10
               Top             =   2460
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               Height          =   375
               Left            =   8340
               TabIndex        =   19
               Top             =   240
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   315
               Left            =   7680
               TabIndex        =   18
               Top             =   270
               Width           =   435
               _Version        =   1048578
               _ExtentX        =   767
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "克重:"
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   360
               TabIndex        =   17
               Top             =   300
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "白坯名称:"
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   4200
               TabIndex        =   16
               Top             =   300
               Width           =   675
               _Version        =   1048578
               _ExtentX        =   1191
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "门幅:"
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   315
               Left            =   360
               TabIndex        =   15
               Top             =   1350
               Width           =   555
               _Version        =   1048578
               _ExtentX        =   979
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "数量:"
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   4200
               TabIndex        =   14
               Top             =   1380
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单价:"
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   7680
               TabIndex        =   13
               Top             =   1380
               Width           =   555
               _Version        =   1048578
               _ExtentX        =   979
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "金额:"
            End
            Begin XtremeSuiteControls.Label Label6 
               Height          =   255
               Left            =   360
               TabIndex        =   12
               Top             =   2520
               Width           =   555
               _Version        =   1048578
               _ExtentX        =   979
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "交期:"
            End
            Begin XtremeSuiteControls.Label Label7 
               Height          =   375
               Left            =   4200
               TabIndex        =   11
               Top             =   2460
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "备注:"
            End
         End
      End
   End
End
Attribute VB_Name = "frmWhitecontract_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As String
Public OriginalProduct As String
Public bsave As Boolean
Public itemid As String
Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "保存"
                save
        Case "退出"
                Unload Me
    End Select
End Sub

Private Sub FlatEdit3_Change()
    FlatEdit5.Text = Val(FlatEdit3.Text) * Val(FlatEdit4.Text)
End Sub

Private Sub FlatEdit4_Change()
    FlatEdit5.Text = Val(FlatEdit3.Text) * Val(FlatEdit4.Text)
End Sub

Private Sub Form_Load()
    InitFrm
    bsave = False
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    DTPicker1.Value = Now
End Sub

Private Sub PushButton1_Click()
        Dim frm1 As New frmpopupWhite
        frm1.Show vbModal
        OriginalProduct = frm1.Whiteid
        FlatEdit1.Text = frm1.WhiteName
        Unload frm1
End Sub

Private Sub save()
    Dim sql As String
    Dim rs As New RecordSet
    If Trim(FlatEdit1.Text) = "" Then
        MsgBox "白坯名称不能空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit2.Text) = "" Then
        MsgBox "门幅不能空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit6.Text) = "" Then
        MsgBox "克重不能空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit3.Text) = "" Then
        MsgBox "数量不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit4.Text) = "" Then
        MsgBox "单价不能为空", vbInformation, "提示"
        Exit Sub
    End If
    sql = "select * from G_Billwhite where B_ID='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If Len(itemid) > 0 Then
            Update_Edit
        Else
            saveALL_Edit
        End If
    Else
        If Len(itemid) > 0 Then
            Update
        Else
            saveALL
        End If
    End If
     bsave = True
     Me.Hide
End Sub

Private Sub Update_Edit()
     If validation = True Then
         Dim sql As String
        Dim a As String
        a = Format(DTPicker1.Value, "YYYY-MM-DD")
        sql = "update G_BilldetailYarn set B_GoodsID='" & OriginalProduct & "',B_width='" & FlatEdit2.Text & "',B_UnitWeight='" & FlatEdit6.Text & "',B_Qty='" & FlatEdit3.Text & "'"
        sql = sql & ",B_price='" & FlatEdit4.Text & "',B_sum='" & FlatEdit5.Text & "',B_DeliveryTime='" & a & "',B_MemoDetail='" & FlatEdit7.Text & "'"
        sql = sql & " where B_itemid='" & itemid & "'"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
    End If
End Sub

Private Sub saveALL_Edit()
    If validation = True Then
        Dim sql As String
        Dim a As String
        Dim sql1 As String
        a = Format(DTPicker1.Value, "YYYY-MM-DD")
        Dim rs As New RecordSet
        sql = "select * from G_draftBilldetailYarn where 1=1"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rs.addnew
        rs!B_datecreate = Now
        rs.Update
        itemid = rs!B_itemid
        sql1 = "exec usp_InsertwhiteOrder '" & itemid & "','" & id & "','" & OriginalProduct & "','" & FlatEdit2.Text & "','" & FlatEdit6.Text & "'"
        sql1 = sql1 & ",'" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit5.Text & "','" & a & "','" & FlatEdit7.Text & "','" & "" & "','" & B_WhiteOrderid & "'"
        Gm.cnnTool.cnn.Execute sql1
    End If
End Sub
Private Sub saveALL()
'    Dim sql As String
'    Dim sql1 As String
'
'    Dim sql2 As String
'    Dim a As String
'    a = Format(DTPicker1.Value, "YYYY-MM-DD")
'    Dim rs As New RecordSet
'    sql = "select * from G_draftBilldetailWhite where 1=1"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'    itemid = rs!B_itemid
'
'    sql1 = "exec usp_Whitecontract '" & itemid & "','" & id & "','" & OriginalProduct & "','" & FlatEdit2.Text & "',"
'    sql1 = sql1 & "'" & FlatEdit6.Text & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit5.Text & "',"
'    sql1 = sql1 & "'" & a & "','" & FlatEdit7.Text & "','" & B_WhiteOrderid & "'"
'    Gm.cnnTool.cnn.Execute sql1
    Dim sql As String
    Dim a As String
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    Dim rs As New RecordSet
    sql = "select * from G_draftBilldetailwhite where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
        rs!B_id = id
        rs!B_GoodsID = OriginalProduct
        rs!B_Width = FlatEdit2.Text
        rs!B_UnitWeight = FlatEdit6.Text
        rs!B_qty = FlatEdit3.Text
        rs!B_price = FlatEdit4.Text
        rs!B_sum = FlatEdit5.Text
        rs!B_DeliveryTime = a
        rs!B_MemoDetail = FlatEdit7.Text
        rs!B_orderid = B_WhiteOrderid
    rs.Update
End Sub

Private Sub Update()
    Dim sql As String
    Dim a As String
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    sql = "update G_BilldetailWhite set B_GoodsID='" & OriginalProduct & "',B_width='" & FlatEdit2.Text & "',B_UnitWeight='" & FlatEdit6.Text & "' ,B_Qty='" & FlatEdit3.Text & "'"
    sql = sql & ",B_price='" & FlatEdit4.Text & "',B_sum='" & FlatEdit5.Text & "',B_DeliveryTime='" & a & "',B_MemoDetail='" & FlatEdit7.Text & "'"
    sql = sql & " where B_itemid='" & itemid & "'"
    Gm.cnnTool.cnn.Execute sql
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
End Sub
Private Sub FlatEdit4_KeyPress(KeyAscii As Integer)
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

Private Function validation() As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_SystemUser where B_username='" & Gm.SysID.SystemUser & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        validation = True
        Exit Function
    End If
    If Gm.SysID.SystemUser = UserName Then
        validation = True
    Else
        validation = False
        MsgBox "不能修改其他人做的数据", vbInformation, "提示"
    End If
End Function
