VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOrderColor 
   Caption         =   "色布采购"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11820
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
   ScaleHeight     =   8010
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   8010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _LayoutVersion  =   1
      _ExtentX        =   20849
      _ExtentY        =   14129
      _DataPath       =   ""
      Bands           =   "frmOrderColor.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6975
         Left            =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   10755
         _cx             =   18971
         _cy             =   12303
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
         _GridInfo       =   $"frmOrderColor.frx":17E8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   4755
            Left            =   90
            ScaleHeight     =   4755
            ScaleWidth      =   10575
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   10575
            Begin VB.PictureBox Picture2 
               Height          =   375
               Left            =   8160
               ScaleHeight     =   315
               ScaleWidth      =   1515
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   1200
               Width           =   1575
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit12 
               Height          =   375
               Left            =   4860
               TabIndex        =   3
               Top             =   2250
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   6420
               TabIndex        =   4
               Top             =   180
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1440
               TabIndex        =   5
               Top             =   180
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
               Enabled         =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4860
               TabIndex        =   6
               Top             =   825
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   8040
               TabIndex        =   7
               Top             =   180
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1440
               TabIndex        =   8
               Top             =   825
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit6 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1440
               TabIndex        =   9
               Top             =   1545
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit8 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   8040
               TabIndex        =   10
               Top             =   1545
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit11 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4860
               TabIndex        =   11
               Top             =   180
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
               Height          =   375
               Left            =   1440
               TabIndex        =   12
               Top             =   2250
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit13 
               Height          =   735
               Left            =   1380
               TabIndex        =   13
               Top             =   3960
               Width           =   8835
               _Version        =   1048578
               _ExtentX        =   15584
               _ExtentY        =   1296
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               MultiLine       =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   8040
               TabIndex        =   24
               Top             =   825
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
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
               Enabled         =   0   'False
               BackColor       =   16777215
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit16 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   4860
               TabIndex        =   26
               Top             =   1545
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
            Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
               Height          =   375
               Left            =   8040
               TabIndex        =   29
               Top             =   2250
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   68
               Format          =   1
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   375
               Left            =   9720
               TabIndex        =   30
               Top             =   825
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit9 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1440
               TabIndex        =   33
               Top             =   2820
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
            Begin XtremeSuiteControls.DateTimePicker DateTimePicker2 
               Height          =   375
               Left            =   8040
               TabIndex        =   35
               Top             =   2880
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   68
               Format          =   1
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   390
               Left            =   4860
               TabIndex        =   37
               Top             =   2880
               Width           =   1935
               _Version        =   1048578
               _ExtentX        =   3413
               _ExtentY        =   688
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit10 
               DataField       =   "B_CodeID"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1440
               TabIndex        =   39
               Top             =   3360
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
            Begin XtremeSuiteControls.Label Label17 
               Height          =   255
               Left            =   360
               TabIndex        =   40
               Top             =   3420
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "金额："
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
            Begin XtremeSuiteControls.Label Label16 
               Height          =   255
               Left            =   3600
               TabIndex        =   38
               Top             =   2895
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "计价单位:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label15 
               Height          =   255
               Left            =   6960
               TabIndex        =   36
               Top             =   2940
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "下单日期："
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
            Begin XtremeSuiteControls.Label Label11 
               Height          =   255
               Left            =   360
               TabIndex        =   34
               Top             =   2880
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单价："
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
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   7080
               TabIndex        =   28
               Top             =   2310
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "交期："
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
            Begin XtremeSuiteControls.Label Label12 
               Height          =   255
               Left            =   3600
               TabIndex        =   27
               Top             =   1605
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "匹数："
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
            Begin XtremeSuiteControls.Label Label6 
               Height          =   255
               Left            =   7080
               TabIndex        =   25
               Top             =   885
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "颜色："
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
               Left            =   240
               TabIndex        =   23
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
               Left            =   3600
               TabIndex        =   22
               Top             =   885
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
               Left            =   7080
               TabIndex        =   21
               Top             =   240
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
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
               Left            =   360
               TabIndex        =   20
               Top             =   885
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   3600
               TabIndex        =   19
               Top             =   240
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "色布供应商："
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
            Begin XtremeSuiteControls.Label Label7 
               Height          =   375
               Left            =   360
               TabIndex        =   18
               Top             =   1545
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "色号："
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
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   360
               TabIndex        =   17
               Top             =   2310
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "米数："
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
               Height          =   255
               Left            =   7080
               TabIndex        =   16
               Top             =   1605
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "公斤数："
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
            Begin XtremeSuiteControls.Label Label13 
               Height          =   255
               Left            =   3600
               TabIndex        =   15
               Top             =   2310
               Width           =   555
               _Version        =   1048578
               _ExtentX        =   979
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "码数:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label14 
               Height          =   315
               Left            =   360
               TabIndex        =   14
               Top             =   4290
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "备注:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   1980
            Left            =   90
            TabIndex        =   31
            Top             =   4905
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   3493
            _LayoutType     =   0
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   953
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).DividerColor=   15790320
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=17"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1.5
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   15790320
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(8)   =   ":id=1,.fontname=宋体"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bgbmp=1,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(12)  =   ":id=2,.fontname=宋体"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgpicMode=2,.bgbmp=2,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(15)  =   ":id=3,.fontname=宋体"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
            _StyleDefs(65)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
            _StyleDefs(66)  =   "bmp(1):id=1,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
            _StyleDefs(67)  =   "bmp(2):id=1,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
            _StyleDefs(68)  =   "bmp(3):id=1,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(69)  =   "bmp(4):id=1,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(70)  =   "bmp(5):id=1,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
            _StyleDefs(71)  =   "bmp(6):id=1,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(72)  =   "bmp(7):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(73)  =   "bmp(8):id=1,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(74)  =   "bmp(9):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(75)  =   "bmp(10):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
            _StyleDefs(76)  =   "bmp(11):id=1,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(77)  =   "bmp(12):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(78)  =   "bmp(13):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
            _StyleDefs(79)  =   "bmp(14):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(80)  =   "bmp(15):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(81)  =   "bmp(16):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
            _StyleDefs(82)  =   "bmp(17):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(83)  =   "bmp(18):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(84)  =   "bmp(19):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
            _StyleDefs(85)  =   "bmp(20):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(86)  =   "bmp(21):id=1,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(87)  =   "bmp(22):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
            _StyleDefs(88)  =   "bmp(23):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(89)  =   "bmp(24):id=1,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(90)  =   "bmp(25):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
            _StyleDefs(91)  =   "bmp(26):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(92)  =   "bmp(27):id=1,797r797r797r7wAAAA=="
            _StyleDefs(93)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
            _StyleDefs(94)  =   "bmp(1):id=2,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
            _StyleDefs(95)  =   "bmp(2):id=2,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
            _StyleDefs(96)  =   "bmp(3):id=2,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(97)  =   "bmp(4):id=2,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(98)  =   "bmp(5):id=2,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
            _StyleDefs(99)  =   "bmp(6):id=2,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(100) =   "bmp(7):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(101) =   "bmp(8):id=2,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(102) =   "bmp(9):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(103) =   "bmp(10):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
            _StyleDefs(104) =   "bmp(11):id=2,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(105) =   "bmp(12):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(106) =   "bmp(13):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
            _StyleDefs(107) =   "bmp(14):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(108) =   "bmp(15):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(109) =   "bmp(16):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
            _StyleDefs(110) =   "bmp(17):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(111) =   "bmp(18):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(112) =   "bmp(19):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
            _StyleDefs(113) =   "bmp(20):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(114) =   "bmp(21):id=2,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(115) =   "bmp(22):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
            _StyleDefs(116) =   "bmp(23):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(117) =   "bmp(24):id=2,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(118) =   "bmp(25):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
            _StyleDefs(119) =   "bmp(26):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(120) =   "bmp(27):id=2,797r797r797r7wAAAA=="
         End
      End
   End
End
Attribute VB_Name = "frmOrderColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private client As String
Private colorid As String
Public id As String
Private B_orderitemid As String
Private rss As RecordSet
Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "保存"
            save
        Case "退出"
            Unload Me
        Case "新增"
            AddNew
        Case "删除"
            de
    End Select
End Sub

Private Sub ComboBox1_Click()
    If ComboBox1.Text = "米数" Then
        FlatEdit10.Text = Val(FlatEdit7.Text) * Val(FlatEdit9.Text)
    End If
    If ComboBox1.Text = "公斤数" Then
        FlatEdit10.Text = Val(FlatEdit8.Text) * Val(FlatEdit9.Text)
    End If
    If ComboBox1.Text = "码数" Then
        FlatEdit10.Text = Val(FlatEdit12.Text) * Val(FlatEdit9.Text)
    End If
End Sub
Private Sub Valuation()
    ComboBox1.AddItem "米数"
    ComboBox1.AddItem "公斤数"
    ComboBox1.AddItem "码数"
End Sub
Private Sub Form_Load()
    InitFrm
    Valuation
    DateTimePicker1.Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
     DateTimePicker2.Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
    Grid
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub



Private Sub PushButton1_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "色布供应商"
     frm1.Show vbModal
    client = frm1.clientid
    FlatEdit11.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub Grid()
    Set rss = New RecordSet
    Dim sql As String
    sql = "exec usp_SelectColororder '" & id & "' "
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    TDBGrid1.DataSource = rss
    If rss.RecordCount > 0 Then
        rss.MoveFirst
    End If
    SetGrid
    
    
    
End Sub
Private Sub SetGrid()
    TDBGrid1.Columns("B_ItemIDB").Caption = "订单号"
    TDBGrid1.Columns("B_ClientName").Caption = "色布供应商"
    TDBGrid1.Columns("B_Width").Caption = "门幅"
    TDBGrid1.Columns("B_weight").Caption = "克重"
    TDBGrid1.Columns("B_GoodsNameAlias").Caption = "品名"
    TDBGrid1.Columns("B_Name").Caption = "颜色"
    TDBGrid1.Columns("B_SeHao").Caption = "色号"
    TDBGrid1.Columns("B_ps").Caption = "匹数"
    TDBGrid1.Columns("B_kg").Caption = "公斤"
    TDBGrid1.Columns("B_meter").Caption = "米数"
    TDBGrid1.Columns("B_qty").Caption = "码数"
    TDBGrid1.Columns("B_departdate").Caption = "交期"
    TDBGrid1.Columns("B_MemoDetail").Caption = "备注"
    TDBGrid1.Columns("B_hex").Caption = "色块"
     TDBGrid1.Columns("B_price").Caption = "单价"
    TDBGrid1.Columns("B_ItemIDB").width = 900
    TDBGrid1.Columns("B_Width").width = 900
    TDBGrid1.Columns("B_weight").width = 900
    TDBGrid1.Columns("B_SeHao").width = 900
    TDBGrid1.Columns("B_ps").width = 900
    TDBGrid1.Columns("B_meter").width = 900
    TDBGrid1.Columns("B_kg").width = 900
    TDBGrid1.Columns("B_hex").width = 900
    TDBGrid1.Columns("B_price").width = 900
    TDBGrid1.Columns("B_qty").width = 900
    
    TDBGrid1.Columns("B_Clientid").Visible = False
    TDBGrid1.Columns("B_Clientid").AllowSizing = False
    TDBGrid1.Columns("B_Clientid").Locked = True
    TDBGrid1.Columns("B_color").Visible = False
    TDBGrid1.Columns("B_color").AllowSizing = False
    TDBGrid1.Columns("B_color").Locked = True
    TDBGrid1.Columns("B_itemid").Visible = False
    TDBGrid1.Columns("B_itemid").AllowSizing = False
    TDBGrid1.Columns("B_itemid").Locked = True
    TDBGrid1.Columns("B_XDDate").Visible = False
    TDBGrid1.Columns("B_XDDate").AllowSizing = False
    TDBGrid1.Columns("B_XDDate").Locked = True
    TDBGrid1.Columns("B_BackMaterial").Visible = False
    TDBGrid1.Columns("B_BackMaterial").AllowSizing = False
    TDBGrid1.Columns("B_BackMaterial").Locked = True
    TDBGrid1.Columns("B_Sum").Visible = False
    TDBGrid1.Columns("B_Sum").AllowSizing = False
    TDBGrid1.Columns("B_Sum").Locked = True
    
    TDBGrid1.Columns("B_Hex").FetchStyle = True
    TDBGrid1.MarqueeStyle = dbgHighlightRow
    TDBGrid1.HoldFields
End Sub

Private Sub save()
    If Trim(client) = "" Then
        MsgBox "色布供应商不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit3.Text) = "" Then
        MsgBox "门幅不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit4.Text) = "" Then
        MsgBox "克重不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit2.Text) = "" Then
        MsgBox "品名不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(colorid) = "" And Trim(FlatEdit6.Text) = "" Then
        MsgBox "颜色或者色号不能为空", vbInformation, "提示"
        Exit Sub
    End If
    
    If Trim(FlatEdit16.Text) = "" And Trim(FlatEdit8.Text) = "" And Trim(FlatEdit7.Text) = "" And Trim(FlatEdit12.Text) = "" Then
        MsgBox "匹数,公斤数,米数,码数,任写其一", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(FlatEdit9.Text) = "" Then
        MsgBox "单价不能为空", vbInformation, "提示"
        Exit Sub
    End If
     If Trim(ComboBox1.Text) = "" Then
            MsgBox "计价单位不能为空", vbInformation, "提示"
            Exit Sub
    End If
    
    If B_orderitemid = "" Then
        saveALL
    Else
        saveupdate
    End If
End Sub
Private Sub saveALL()
    Dim idd As String
    Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    Dim sql2 As String
    
    sql = "select * from G_draftBilldetailColor where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    idd = rs!B_ItemID
    
    sql1 = "select * from G_BilldetailColor where 1=1"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs1.AddNew
    rs1!B_ItemID = idd
    rs1!B_ItemIDB = FlatEdit1.Text
    rs1!B_Clientid = client
    rs1!B_Width = FlatEdit3.Text
    rs1!B_weight = FlatEdit4.Text
    rs1!B_GoodsNameAlias = FlatEdit2.Text
    rs1!B_color = colorid
    rs1!B_orderColor = FlatEdit5.Text
    rs1!B_SeHao = FlatEdit6.Text
    rs1!B_price = FlatEdit9.Text
    If Len(Trim(FlatEdit16.Text)) > 0 Then
        rs1!B_ps = IIf(IsNull(FlatEdit16.Text), 0, FlatEdit16.Text)
    End If
    
    
      If Len(Trim(FlatEdit8.Text)) > 0 Then
       rs1!B_kg = IIf(IsNull(FlatEdit8.Text), 0, FlatEdit8.Text)
    End If
    If Len(Trim(FlatEdit7.Text)) > 0 Then
        rs1!B_meter = IIf(IsNull(FlatEdit7.Text), 0, FlatEdit7.Text)
    End If
    If Len(Trim(FlatEdit12.Text)) > 0 Then
        rs1!B_qty = IIf(IsNull(FlatEdit12.Text), 0, FlatEdit12.Text)
    End If
    rs1!B_departdate = DateTimePicker1.Value
    
    rs1!B_XDDate = DateTimePicker2.Value '下单时间
    rs1!B_Sum = IIf(IsNull(FlatEdit10.Text), 0, FlatEdit10.Text) '金额
    rs1!B_BackMaterial = IIf(IsNull(ComboBox1.Text), 0, ComboBox1.Text)
    
    rs1!B_MemoDetail = FlatEdit13.Text
    rs1!B_intype = 1
    
    rs1!B_orderitemid = id
    rs1.Update
    
    sql2 = "delete from G_draftBilldetailColor where B_itemid='" & idd & "'"
    Gm.cnnTool.cnn.Execute sql2
    MsgBox "保存成功", vbInformation, "提示"
    rss.requery
    
End Sub

Private Sub PushButton2_Click()
'    Dim frm1 As New frmpopupColor
'    frm1.Show vbModal
'    If frm1.bsaved = True Then
'        colorid = frm1.colorid
'        FlatEdit5.Text = frm1.colorname
'    End If
    
    
    On Error Resume Next
    Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmpopupColor
    frm1.Show vbModal
    FlatEdit5.Text = Trim(frm1.colorname)
    colorid = frm1.colorid
    If frm1.bsaved = True Then
        FlatEdit5.Enabled = True
    Else
        If colorid = "" Then
             FlatEdit5.Enabled = False
        End If
    End If
        sql = "select * from G_Color where B_SID='" & colorid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        If rs!B_hex <> "" Then
        
            Picture2.BackColor = rs!B_hex
        Else
            Picture2.BackColor = &H8000000F
        End If
    
        
    Unload frm1
    
    
End Sub

Private Sub saveupdate()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "exec usp_OrderColor '" & B_orderitemid & "','" & FlatEdit1.Text & "','" & client & "','" & FlatEdit3.Text & "','" & FlatEdit4.Text & "','" & FlatEdit2.Text & "',"
    sql = sql & "'" & colorid & "','" & FlatEdit6.Text & "','" & FlatEdit16.Text & "','" & FlatEdit8.Text & "',"
    sql = sql & "'" & FlatEdit7.Text & "','" & FlatEdit12.Text & "','" & DateTimePicker1.Value & "','" & FlatEdit13.Text & "','" & FlatEdit5.Text & "','" & FlatEdit9.Text & "','" & FlatEdit10.Text & "','" & ComboBox1.Text & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset
    MsgBox "保存完成", vbInformation, "提示"
    rss.requery
End Sub

Private Function ClientName(ByVal client As String)
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_ContactCompany where B_clientid='" & client & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        ClientName = rs!B_ClientName
    Else
        ClientName = ""
    End If
End Function

Private Function colorname(ByVal client As String)
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_Color where B_sid='" & client & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        colorname = rs!B_name
    Else
        colorname = ""
    End If
End Function

Private Sub FlatEdit16_KeyPress(KeyAscii As Integer)
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

Private Sub FlatEdit8_KeyPress(KeyAscii As Integer)
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

Private Sub FlatEdit12_KeyPress(KeyAscii As Integer)
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
Private Sub FlatEdit9_KeyPress(KeyAscii As Integer)
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



Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    B_orderitemid = rss!B_ItemID
    FlatEdit1.Text = rss!B_ItemIDB
    client = rss!B_Clientid
    FlatEdit11.Text = rss!B_ClientName
    FlatEdit3.Text = rss!B_Width
    FlatEdit4.Text = rss!B_weight
    FlatEdit2.Text = rss!B_GoodsNameAlias
    colorid = rss!B_color
    FlatEdit5.Text = rss!B_name
    FlatEdit6.Text = rss!B_SeHao
    FlatEdit16.Text = rss!B_ps
    FlatEdit8.Text = rss!B_kg
    FlatEdit7.Text = IIf(IsNull(rss!B_meter), "", rss!B_meter)
    FlatEdit12.Text = IIf(IsNull(rss!B_qty), "", rss!B_qty)
    DateTimePicker1.Value = rss!B_departdate
    DateTimePicker2.Value = rss!B_XDDate
    FlatEdit13.Text = rss!B_MemoDetail
    Picture2.BackColor = rss!B_hex
    FlatEdit9.Text = rss!B_price
    FlatEdit10.Text = IIf(IsNull(rrss!B_Sum), "", rss!B_Sum)
    ComboBox1.Text = IIf(IsNull(rss!B_BackMaterial), "", rss!B_BackMaterial)
    
End Sub
Private Sub AddNew()
    Dim a As String
    a = FlatEdit1.Text
    B_orderitemid = ""
    colorid = ""
    client = ""
    Dim o As Object
    
    For Each o In Me.Controls
        Select Case TypeName(o)
        
            Case "FlatEdit"
                o.Text = ""
            Case "ComboBox"
                o.Text = ""
        End Select
    Next
    FlatEdit1.Text = a
End Sub

Private Sub de()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    sql = "delete from G_BilldetailColor where B_itemid='" & rss!B_ItemID & "'" '
    Gm.cnnTool.cnn.Execute sql
    rss.requery
End Sub
Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
'
'    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid1.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
End Sub

