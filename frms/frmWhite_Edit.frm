VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmWhite_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "白坯计划"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWhite_Edit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   11445
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   4755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11445
      _LayoutVersion  =   1
      _ExtentX        =   20188
      _ExtentY        =   8387
      _DataPath       =   ""
      Bands           =   "frmWhite_Edit.frx":038A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   4095
         Left            =   360
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   10815
         _cx             =   19076
         _cy             =   7223
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
         BorderWidth     =   6
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
         GridRows        =   4
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmWhite_Edit.frx":1136
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   3915
            Left            =   90
            ScaleHeight     =   3915
            ScaleWidth      =   10635
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   10635
            Begin VB.CheckBox Check1 
               Caption         =   "是否采购"
               Height          =   375
               Left            =   360
               TabIndex        =   25
               Top             =   3360
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   345
               Left            =   4680
               TabIndex        =   21
               Top             =   2400
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   609
               _Version        =   393216
               Format          =   222167041
               CurrentDate     =   43099
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   9600
               TabIndex        =   14
               Top             =   300
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4680
               TabIndex        =   3
               Top             =   300
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16777215
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   8160
               TabIndex        =   8
               Top             =   300
               Width           =   1455
               _Version        =   1048578
               _ExtentX        =   2566
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1320
               TabIndex        =   11
               Top             =   2400
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   345
               Left            =   1320
               TabIndex        =   15
               Top             =   330
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   609
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox ComboBox2 
               Height          =   345
               Left            =   4680
               TabIndex        =   16
               Top             =   1320
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   609
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.ComboBox ComboBox3 
               Height          =   345
               Left            =   8160
               TabIndex        =   17
               Top             =   1320
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   609
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1320
               TabIndex        =   19
               Top             =   1320
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16777215
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   375
               Left            =   6180
               TabIndex        =   22
               Top             =   3360
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit11 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4680
               TabIndex        =   23
               Top             =   3360
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BackColor       =   14737632
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit7 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   1455
               Left            =   8160
               TabIndex        =   13
               Top             =   2400
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   2566
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   3480
               TabIndex        =   24
               Top             =   3420
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "白坯供应商："
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   375
               Left            =   3840
               TabIndex        =   20
               Top             =   2400
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "交期："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   375
               Left            =   360
               TabIndex        =   18
               Top             =   1320
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "毛高:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label7 
               Height          =   375
               Left            =   6960
               TabIndex        =   12
               Top             =   2400
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "备注："
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label6 
               Height          =   495
               Left            =   360
               TabIndex        =   10
               Top             =   2340
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "数量kg："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   6960
               TabIndex        =   9
               Top             =   360
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "白坯名称："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   360
               TabIndex        =   7
               Top             =   360
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   3840
               TabIndex        =   6
               Top             =   360
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "品名："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Left            =   3840
               TabIndex        =   5
               Top             =   1380
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "门幅："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   6960
               TabIndex        =   4
               Top             =   1380
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "克重："
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
End
Attribute VB_Name = "frmWhite_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As String
Public itemid As String
Public Whiteid As String
Public theidwhite As String
Private clsBL As clsBL
Private clsCLOrderID As New clsCJComboLinker

Private theOrderID As String
Private strSQL As String
Public client As String


Public Property Let OrderID(ByVal vData As String)
    theOrderID = vData
End Property

Private Sub FillData()
    If Len(theOrderID) <= 0 Then
        Exit Sub
    End If
    
    SetComboListIndex theOrderID
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "保存"
            save
        Case "退出"
            Unload Me
    
    End Select
End Sub

Private Sub save()
       If yanzhenWhite(id) = False Then
                Exit Sub
        End If
      If Trim(ComboBox1.Text) = "" Then
        MsgBox "订单号不能为空", vbInformation, "提示"
        Exit Sub
    End If
      If Trim(FlatEdit2.Text) = "" Then
        MsgBox "品名不能为空", vbInformation, "提示"
        Exit Sub
    End If
      If Trim(ComboBox2.Text) = "" Then
        MsgBox "门幅不能为空", vbInformation, "提示"
        Exit Sub
    End If
      If Trim(ComboBox3.Text) = "" Then
        MsgBox "克重不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit5.Text) = "" Then
        MsgBox "白坯名称不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit6.Text) = "" Then
        MsgBox "数量不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Check1.Value = 1 Then
        If Trim(client) = "" Then
            MsgBox "供应商不能为空", vbInformation, "提示"
            Exit Sub
        End If
    End If
  
    Savedetailwhite
    createdate (id)
    Me.Hide
End Sub

Private Sub Savedetailwhite()
     If Len(itemid) > 0 Then
        savedetail_update
    Else
'        Dim rs1 As New RecordSet
'        Dim sql1 As String
'        sql1 = "select *from G_BillWhite where B_id='" & theidwhite & "'"
'        rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'        If rs1.RecordCount > 0 Then
'            Detail
'        Else
'            Dim rs3 As New RecordSet
'            Dim sql3 As String
'            sql3 = "exec usp_savedetailwhite '" & theidwhite & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "','" & Whiteid & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "'"
'            Debug.Print sql3
'            rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'         End If
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BillWhite where B_belongorderid='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
           savemain
    Else
            theidwhite = rs!B_id
'           SaveDetail
    End If
    savedetail
 End If
    Me.Hide
End Sub
Private Sub savemain()
    Set clsBL = New clsBL
    Dim sql As String
            Dim rs As New RecordSet
            sql = "select * from G_DraftBillWhite where 1=1 "
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs.AddNew
            Dim a As String
            a = Format(Now, "YYYY-MM-DD")
            rs!B_datecreate = a
            rs.Update
            theidwhite = rs!B_id
            
            Dim rs1 As New RecordSet
            Dim sql1 As String
            sql1 = "select * from G_BillWhite where 1=1 "
            rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs1.AddNew
               rs1!B_id = theidwhite
               Dim b As String
               b = Format(Now, "YYYY-MM-DD")
               rs1!B_datecreate = b
               rs1!B_BID = B_BID
               rs1!B_ObjectID = B_ObjectID
               rs1!B_BillType = B_BillType
               rs1!B_UserName = Gm.SysID.SystemUser
               rs1!B_Codeid = clsBL.GetFrameCodeDetail_01(B_ObjectID)
               rs1!B_BelongOrderID = id
               rs1.Update
               Dim rs2 As New RecordSet
               Dim sql2 As String
               sql2 = "delete from G_DraftBillWhite where B_ID='" & theidwhite & "'"
               Gm.cnnTool.cnn.Execute sql2
     
End Sub

Private Sub savedetail()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_DraftBillDetailWhite where 1=0"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Dim a As String
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    
     itemid = rs!B_ItemID
     
     Dim lIncr As Long
     Dim szBC13 As String
     lIncr = GetNewBCIncr
     szBC13 = GetBC13(FillGetBC12(lIncr))
     Debug.Print szBC13
     Dim rs1 As New RecordSet
     Dim sql1 As String
     sql1 = "select * from G_BillDetailWhite where 1=0"
      rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
      rs1.AddNew
      rs1!B_ItemID = itemid
      rs1!B_id = theidwhite
      rs1!B_ItemIDB = ComboBox1.Text
      rs1!B_GoodsNameAlias = FlatEdit2.Text
      rs1!B_GoodsID = Whiteid
      rs1!B_Width = ComboBox2.Text
      rs1!B_UnitWeight = ComboBox3.Text
      rs1!B_BoxQty = FlatEdit6.Text
      rs1!B_MemoDetail = FlatEdit7.Text
      rs1!B_Maohight = FlatEdit1.Text
      rs1!B_BCIncr = lIncr
      rs1!B_BC13 = szBC13
      rs1!B_Deliverydate = a
      rs1!B_intype = Check1.Value
      If Check1.Value = 1 Then
        rs1!B_supplier = client
      End If
      rs1.Update
      
      Dim sql2 As String
      sql2 = "delete from G_DraftBillDetailWhite where B_itemid='" & itemid & "'"
      Gm.cnnTool.cnn.Execute sql2
End Sub

Private Sub ComboBox1_Click()
    If Len(ComboBox1.Text) > 0 Then
        Dim rs2 As New RecordSet
        Dim sql2 As String
        sql2 = "select  CASE WHEN LEN(ISNULL(b.B_Name,0))>1 THEN b.B_Name ELSE a.B_GoodsID  END AS B_GoodsID  from G_billdetailorder a LEFT OUTER JOIN G_Product b ON a.B_GoodsID=b.B_SID where a.B_OrderCode='" & ComboBox1.Text & "' and a.B_ID='" & id & "'  GROUP BY a.B_GoodsID,b.B_Name"
        rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
       Debug.Print sql2
'        FlatEdit2.Text = IIf(IsNull(rs2!B_name), "", rs2!B_name)
        FlatEdit2.Text = IIf(IsNull(rs2!B_GoodsID), "", rs2!B_GoodsID)
        
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select distinct B_Width from G_billdetailorder where B_OrderCode='" & ComboBox1.Text & "' and B_ID='" & id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "select distinct B_Weight from G_billdetailorder where B_OrderCode='" & ComboBox1.Text & "' and B_ID='" & id & "'"
        rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        rs.MoveFirst
        rs1.MoveFirst
        ComboBox2.Clear
        ComboBox3.Clear
        Do While Not rs.EOF
            ComboBox2.AddItem rs!B_Width
            rs.movenext
        Loop
        Do While Not rs1.EOF
            ComboBox3.AddItem rs1!B_weight
            rs1.movenext
        Loop
        If rs.RecordCount = 1 Then
             ComboBox2.ListIndex = 0
        End If
       If rs1.RecordCount = 1 Then
            ComboBox3.ListIndex = 0
        End If
    End If
End Sub



Private Sub Form_Load()
    InitFrm
    '绑定订单号
    dingdanhao
    
    
    FillData
End Sub

Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    DTPicker1.Value = Now
End Sub

Private Sub SetComboListIndex(ByVal vData As String)
    Dim i As Long
    For i = 0 To ComboBox1.ListCount - 1
        If ComboBox1.List(i) = vData Then
            ComboBox1.ListIndex = i
        End If
    Next
End Sub

Private Sub dingdanhao()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select distinct B_OrderCode from G_Billdetailorder where B_ID='" & id & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        ComboBox1.AddItem rs!B_ordercode
        rs.movenext
    Loop

'    Set clsCLOrderID = New clsCJComboLinker
'    clsCLOrderID.InitCls ComboBox1, "B_OrderCode", "B_OrderCode", rs

End Sub
Private Sub savedetail_update()
      
        Dim sql2 As String
         sql2 = "exec usp_savebilletailupdate '" & itemid & "','" & ComboBox1.Text & "','" & FlatEdit2.Text & "','" & Whiteid & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "','" & FlatEdit1.Text & "','" & DTPicker1.Value & "','" & Check1.Value & "','" & client & "'"
         Debug.Print sql2
        Gm.cnnTool.cnn.Execute sql2
'    Else
'        Dim sql As String
'        sql = "exec usp_updatedetailwhite '" & itemid & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "','" & Whiteid & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "'"
'        Gm.cnnTool.cnn.Execute sql
'    End If
End Sub




Private Sub PushButton1_Click()
     Dim frm1 As New frmpopupWhite
    frm1.Show vbModal
    FlatEdit5.Text = Trim(frm1.WhiteName)
    Whiteid = frm1.Whiteid
    Unload frm1
End Sub

Private Sub FlatEdit1_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit2.SetFocus
    End Select
End Sub
Private Sub FlatEdit2_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            If Len(FlatEdit5.Text) <= 0 Then
                PushButton1_Click
            Else
                FlatEdit3.SetFocus
            End If
    End Select
End Sub
Private Sub FlatEdit3_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
                FlatEdit4.SetFocus
    End Select
End Sub
Private Sub FlatEdit4_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
                FlatEdit6.SetFocus
    End Select
End Sub
Private Sub FlatEdit6_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
                FlatEdit7.SetFocus
    End Select
End Sub

'从表G_BillDetailColor获取当前最新一个条码的自增数字
Private Function GetNewBCIncr() As Long
    Dim rs As New RecordSet
    strSQL = "select top 1 * from G_BillDetailwhite order by B_BCIncr desc"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Dim lRturn As Long
    If rs.RecordCount <= 0 Then
        lRturn = 1
    Else
        lRturn = IIf(IsNull(rs!B_BCIncr), 0, rs!B_BCIncr) + 1
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetNewBCIncr = lRturn
End Function

'传入参数：任意长度的自增数字的字符串类型
'返回值：返回BC13条码的前面12位字符
Private Function FillGetBC12(ByVal vIncr As String) As String
    Dim cls1 As New clsString
    Dim szReturn As String
    
    szReturn = cls1.FillRepeat(vIncr, 11, "0", True)
    szReturn = COLORBC13FIRST & szReturn
    
    FillGetBC12 = szReturn
End Function

Private Function GetBC13(ByVal vBC12 As String) As String
    Dim szRturn As String
    szRturn = GetEAN13CheckOut(vBC12)
    
    GetBC13 = vBC12 & szRturn
End Function

'获取最新的一个13位条码
Private Function GetBC13Ex() As String
    Dim szIncr As String
    szIncr = GetNewBCIncr
    
    Dim szBC12 As String
    szBC12 = FillGetBC12(GetNewBCIncr)
    
    GetBC13Ex = GetBC13(szBC12)
End Function

'当订单号合同中订单号都有开始计算时间
Private Sub createdate(ByVal id As String)
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "select distinct B_itemidb from G_BilldetailWhite where B_id=(select B_ID from G_BillWhite where B_belongorderid='" & id & "')"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql1 = "select distinct B_orderCode from G_Billdetailorder where B_id='" & id & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql2 = "select * from G_BilldetailWhite where B_id=(select B_ID from G_BillWhite where B_belongorderid='" & id & "')"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount = rs1.RecordCount Then
        Do While Not rs2.EOF
            rs2!B_Date = Now
            rs2.movenext
        Loop
    End If
End Sub
Public Function yanzhenWhite(ByVal theid As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        yanzhenWhite = True
        Exit Function
    End If
    
    sql1 = "select * from G_BilldetailWhite where B_ID=(select B_ID from G_BillWhite where B_belongorderid='" & theid & "')"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    Debug.Print sql1
'    sql2 = "select * from G_Billorder  where B_ID='" & theid & "'"
    sql2 = "SELECT * FROM G_UserPro WHERE B_username='" & Gm.SysID.SystemUser & "' AND B_objectid='11S006'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs2.RecordCount > 0 Then
        If IIf(IsNull(rs2!B_new), 0, rs2!B_new) = 1 Then
            yanzhenWhite = True
        Else
            yanzhenWhite = False
            MsgBox "请设置权限", vbInformation, "提示"
            Exit Function
        End If
        If rs1.RecordCount > 0 Then
            If IIf(IsNull(rs1!B_Date), "", rs1!B_Date) <> "" Then
                    If DateDiff("s", rs1!B_Date, Now) > 84600 Then
                        yanzhenWhite = False
                        MsgBox "已经超过制作本单据的时间不能进行修改", vbInformation, "提示"
                    Else
                        yanzhenWhite = True
                    End If
            End If
        End If
    Else
        yanzhenWhite = False
        MsgBox "你没有此权限", vbInformation, "提示"
    End If
End Function

Private Sub FlatEdit6_KeyPress(KeyAscii As Integer)
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

Private Sub PushButton2_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "白坯供应商"
     frm1.Show vbModal
    client = frm1.clientid
    FlatEdit11.Text = frm1.ClientName
    Unload frm1
End Sub
