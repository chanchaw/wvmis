VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmWhiteOrder 
   Caption         =   "白坯定单"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   8190
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      _LayoutVersion  =   1
      _ExtentX        =   21034
      _ExtentY        =   14446
      _DataPath       =   ""
      Bands           =   "frmWhiteOrder.frx":0000
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
         _GridInfo       =   $"frmWhiteOrder.frx":0C52
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1305
            Left            =   45
            ScaleHeight     =   1305
            ScaleWidth      =   11970
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   45
            Width           =   11970
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   1140
               TabIndex        =   17
               Top             =   660
               Width           =   1755
               _Version        =   1048578
               _ExtentX        =   3096
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
               Left            =   5280
               TabIndex        =   18
               Top             =   720
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
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
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   300
               Left            =   9540
               TabIndex        =   19
               Top             =   720
               Width           =   1575
               _Version        =   1048578
               _ExtentX        =   2778
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2880
               TabIndex        =   20
               Top             =   660
               Width           =   1155
               _Version        =   1048578
               _ExtentX        =   2037
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "清空供应商"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   435
               Left            =   4140
               TabIndex        =   24
               Top             =   120
               Width           =   2955
               _Version        =   1048578
               _ExtentX        =   5212
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "凯鑫白坯申请执行单"
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
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   480
               TabIndex        =   23
               Top             =   780
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "供应商:"
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
               Height          =   315
               Left            =   4530
               TabIndex        =   22
               Top             =   750
               Width           =   675
               _Version        =   1048578
               _ExtentX        =   1191
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "合同号:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   195
               Left            =   8640
               TabIndex        =   21
               Top             =   810
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "是否执行:"
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
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1620
            Left            =   45
            ScaleHeight     =   1620
            ScaleWidth      =   11970
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   3255
            Width           =   11970
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   5460
               TabIndex        =   7
               Top             =   600
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Format          =   204668929
               CurrentDate     =   43059
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   1260
               TabIndex        =   8
               Top             =   563
               Width           =   1635
               _Version        =   1048578
               _ExtentX        =   2884
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
            Begin XtremeSuiteControls.ComboBox ComboBox2 
               Height          =   300
               Left            =   1260
               TabIndex        =   9
               Top             =   1200
               Width           =   1635
               _Version        =   1048578
               _ExtentX        =   2884
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox ComboBox3 
               Height          =   300
               Left            =   5460
               TabIndex        =   10
               Top             =   1200
               Width           =   1635
               _Version        =   1048578
               _ExtentX        =   2884
               _ExtentY        =   529
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               Height          =   795
               Left            =   9540
               TabIndex        =   27
               Top             =   540
               Width           =   2175
               _Version        =   1048578
               _ExtentX        =   3836
               _ExtentY        =   1402
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               MultiLine       =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   8700
               TabIndex        =   28
               Top             =   600
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "备注:"
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   435
               Left            =   4140
               TabIndex        =   15
               Top             =   60
               Width           =   2595
               _Version        =   1048578
               _ExtentX        =   4577
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "凯鑫白坯采购定单"
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
            Begin XtremeSuiteControls.Label Label6 
               Height          =   255
               Left            =   420
               TabIndex        =   14
               Top             =   623
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "供应商:"
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
            Begin XtremeSuiteControls.Label Label7 
               Height          =   195
               Left            =   4260
               TabIndex        =   13
               Top             =   653
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "定单日期:"
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   255
               Left            =   360
               TabIndex        =   12
               Top             =   1200
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "交货方式:"
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   315
               Left            =   4260
               TabIndex        =   11
               Top             =   1193
               Width           =   1155
               _Version        =   1048578
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "运费结算方式:"
            End
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   660
            Left            =   45
            ScaleHeight     =   660
            ScaleWidth      =   11970
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   6825
            Width           =   11970
            Begin XtremeSuiteControls.PushButton PushButton4 
               Height          =   495
               Left            =   9960
               TabIndex        =   3
               Top             =   60
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "退出"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton3 
               Height          =   495
               Left            =   8040
               TabIndex        =   4
               Top             =   60
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "保存并打印"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   495
               Left            =   6120
               TabIndex        =   5
               Top             =   60
               Width           =   1215
               _Version        =   1048578
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "保存"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   1785
            Left            =   45
            TabIndex        =   25
            Top             =   1410
            Width           =   11970
            _ExtentX        =   21114
            _ExtentY        =   3149
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
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
         Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
            Height          =   1830
            Left            =   45
            TabIndex        =   26
            Top             =   4935
            Width           =   11970
            _ExtentX        =   21114
            _ExtentY        =   3228
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
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
Attribute VB_Name = "frmWhiteOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mvarObjectID As String
Private client As String
Private rss As RecordSet
Private rss1 As RecordSet
Private a As String
Private id As String
Private theBLTool As New clsAutoCreateBL
Private Const theObjectID As String = "12B006"

Private Function GetCodeID() As String
    GetCodeID = theBLTool.GetFrameCodeDetail(theObjectID)
End Function

Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property


Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
        Select Case Tool.name
            Case "查询"
                grid
            Case "修改供应商"
                UpdateClient
            Case "选择单行"
                copyOne
            Case "选择全部"
                copyAll
        End Select
End Sub

Private Sub PushButton1_Click()
    Dim sql As String
    If TDBGrid2.ApproxCount > 0 Then
        If MsgBox("采购定单已经有数据,更换客户将删除数据", vbYesNo + vbDefaultButton2 + vbInformation, "提示") = vbNo Then
        
        Else
            sql = "delete from G_DraftBillDetailWhite where B_ID='" & id & "'"
            Gm.cnnTool.cnn.Execute sql
            rss1.requery
                FlatEdit1.Text = ""
                FlatEdit3.Text = ""
                client = ""
        End If
    Else
            FlatEdit1.Text = ""
            FlatEdit3.Text = ""
            client = ""
    End If
    grid
End Sub

Private Sub PushButton4_Click()
    Dim sql As String
    If TDBGrid2.ApproxCount > 0 Then
        If MsgBox("是为保存数据，否为采购定单中数据将会删除", vbYesNo + vbDefaultButton2 + vbInformation, "提示") = vbNo Then
            sql = "delete from G_DraftBillDetailWhite where B_ID='" & id & "'"
            Gm.cnnTool.cnn.Execute sql
'            rss1.requery
            FlatEdit1.Text = ""
            FlatEdit3.Text = ""
            client = ""
            Unload Me
        Else
            PushButton2_Click
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub TDBGrid1_DblClick()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    If Len(FlatEdit1.Text) <= 0 Then
        client = rss!B_Clientid
        FlatEdit1.Text = rss!B_ClientName
        FlatEdit3.Text = rss!B_ClientName
        grid
    Else
        If client <> rss!B_Clientid Then
            MsgBox "供应商中已经有数据，请先清除数据", vbInformation, "提示"
        End If
        Exit Sub
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar21.Bands("网格右键").PopupMenu
    End If
    
End Sub
Private Sub Form_Load()
    InitFrm
    execution
    delivery
    ClearWay
    client = ""
    DTPicker1.Value = Now
    grid
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub
'是否执行
Private Sub execution()
    ComboBox1.AddItem "未执行"
    ComboBox1.AddItem "已执行"
    ComboBox1.AddItem "全部"
    ComboBox1.Text = "未执行"
End Sub
'交货方式
Private Sub delivery()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "Select B_SID From G_Delivery Where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            ComboBox2.AddItem "" & rs!B_sid & ""
            rs.movenext
        Loop
    End If
End Sub
'运费结算方式
Private Sub ClearWay()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "Select B_SID From G_Balance Where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            ComboBox3.AddItem "" & rs!B_sid & ""
            rs.movenext
        Loop
    End If
End Sub
'选择客户
'Private Sub PushButton1_Click()
'    Dim sql As String
'    If TDBGrid2.ApproxCount > 0 Then
'        If MsgBox("修改客户,下面列表将进行删除", vbYesNo + vbDefaultButton2 + vbInformation, "提示") = vbNo Then
'                Exit Sub
'        Else
'            sql = "delete from G_DraftOriginalDetailOrder where B_ID='" & id & "'"
'            Gm.cnnTool.cnn.Execute sql
'            rss1.requery
'        End If
'    End If
'        Dim frm1 As New frmPopupClient_Edit
'         frm1.a = "原料供应商"
'         frm1.Show vbModal
'        Client = frm1.Clientid
'        FlatEdit1.Text = frm1.ClientName
'        Unload frm1
'End Sub

Private Sub grid()
'    If FlatEdit1.Text = "" Then
'        MsgBox "请先选择客户", vbInformation, "提示"
'        Exit Sub
'    End If
    Dim sql As String
    Set rss = New RecordSet
    choose
    Dim f As String
    f = "白坯"
    sql = "exec usp_selectWhiteOrder '" & f & "','" & client & "','" & Trim(FlatEdit2.Text) & "','" & a & "'"
    Debug.Print sql
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid1.DataSource = rss
    setgrid
    sumall
    If rss.RecordCount > 0 Then
        rss.MoveFirst
    End If
End Sub
'下拉列表
Private Sub choose()
    If ComboBox1.Text = "未执行" Then
        a = 0
    End If
    If ComboBox1.Text = "已执行" Then
        a = 1
    End If
    If ComboBox1.Text = "全部" Then
        a = 2
    End If
End Sub
Private Sub setgrid()
    TDBGrid1.Columns("B_ClientName").Caption = "供应商"
    TDBGrid1.Columns("B_pactCode").Caption = "合同号"
    TDBGrid1.Columns("B_itemidb").Caption = "定单号"
    TDBGrid1.Columns("B_GoodsNameAlias").Caption = "白坯名称"
     TDBGrid1.Columns("B_Name").Caption = "原料名称"
    TDBGrid1.Columns("B_width").Caption = "门幅"
    TDBGrid1.Columns("B_UnitWeight").Caption = "克重"
    TDBGrid1.Columns("B_BreedNum").Caption = "数量"
    TDBGrid1.Columns("B_itemid").width = 0
    TDBGrid1.Columns("B_itemid").Visible = False
    TDBGrid1.Columns("B_itemid").AllowSizing = False
    TDBGrid1.Columns("B_Clientid").width = 0
    TDBGrid1.Columns("B_Clientid").Visible = False
    TDBGrid1.Columns("B_Clientid").AllowSizing = False
    TDBGrid1.Columns("B_GoodsNameAlias").width = 0
    TDBGrid1.Columns("B_GoodsNameAlias").Visible = False
    TDBGrid1.Columns("B_GoodsNameAlias").AllowSizing = False
    TDBGrid1.Columns("B_price").width = 0
    TDBGrid1.Columns("B_price").Visible = False
    TDBGrid1.Columns("B_price").AllowSizing = False
    TDBGrid1.MarqueeStyle = dbgHighlightRow
    TDBGrid1.HoldFields
End Sub
'表1页脚小计
Private Sub sumall()
    Dim a As Long
    Dim b As Long
    If rss.RecordCount <= 0 Then
        Exit Sub
    End If

    b = 0
    rss.MoveFirst
    Do While Not rss.EOF
        b = b + IIf(IsNull(rss!B_BreedNum), 0, rss!B_BreedNum)
        rss.movenext
    Loop
    TDBGrid1.Columns("B_ClientName").FooterText = "合计"
  
    TDBGrid1.Columns("B_BreedNum").FooterText = "" & b & ""
End Sub

Private Sub UpdateClient()
    Dim b As String
    Dim sql As String
    Dim sql1 As String
    If TDBGrid1.ApproxCount > 0 Then
        Dim frm1 As New frmPopupClient_Edit
        frm1.a = "白坯加工商"
        frm1.Show vbModal
        If frm1.bsaved = True Then
            b = frm1.clientid
            sql = "update G_WhiteComposition set B_Suppliers='" & b & "' where B_itemid='" & rss!B_itemid & "'"
            Gm.cnnTool.cnn.Execute sql
            If rss!B_Clientid <> b Then
                If TDBGrid2.ApproxCount > 0 Then
                    sql1 = "delete from G_DraftBillDetailWhite where B_ID='" & id & "' and B_orderid='" & rss!B_itemid & "'"
                    Gm.cnnTool.cnn.Execute sql1
                    rss1.requery
                End If
            End If
        End If
     
        Unload frm1
        rss.requery
        sumall
    End If
    rss.MoveFirst
End Sub

'Private Sub TDBGrid1_DblClick()
'    If TDBGrid1.ApproxCount > 0 Then
'        If rss.RecordCount > 0 Then
''            Dim sql As String
''            Dim rs As New RecordSet
''            sql = "select * from G_OriginalOrder where B_Belongorderid"
'            copytodingdan
'            copytodetail
'        End If
'    End If
'End Sub

'复制单行
Private Sub copyOne()
    If TDBGrid1.ApproxCount > 0 Then
        If FlatEdit1.Text = "" Then
            Exit Sub
        Else
            If client <> rss!B_Clientid Then
                MsgBox "供应商不一致", vbInformation, "提示"
                Exit Sub
            End If
        End If
        
    Else
        Exit Sub
    End If

    Dim sql As String
    Dim rs As New RecordSet
    If TDBGrid2.ApproxCount > 0 Then
        sql = "select * from G_DraftBillDetailWhite where B_id='" & id & "' and B_orderid='" & rss!B_itemid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            MsgBox "已经存在此信息", vbInformation, "提示"
            Exit Sub
        End If
        
        copytoOne
    Else
        copytodingdan
        copytoOne
    End If
End Sub
Private Sub copytoOne()
        Dim sql As String
        Dim sql1 As String
        Dim rs As New RecordSet
        sql = "select * from G_DraftBillDetailWhite where 1=1"
        
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        rs.addnew
        rs!B_datecreate = Now
        rs!B_id = id
        rs!B_GoodsID = rss!B_GoodsNameAlias
        rs!B_Width = rss!B_Width
        rs!B_UnitWeight = rss!B_UnitWeight
        rs!B_qty = rss!B_BreedNum
'        rs!B_price = rss!B_price
        rs!B_PactCode = rss!B_PactCode
        rs!B_ItemIDB = rss!B_ItemIDB
        rs!B_orderid = rss!B_itemid
        rs.Update
        Set rss1 = New RecordSet
        sql1 = "exec usp_selectWhitedetail_Edit '" & id & "'"
        Debug.Print sql1
        rss1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        TDBGrid2.DataSource = rss1
        setgrid2
End Sub
'复制全部
Private Sub copyAll()
      Dim sql As String
      Dim rs As RecordSet
      If TDBGrid1.ApproxCount > 0 Then
        If FlatEdit1.Text = "" Then
            Exit Sub
        Else
            Dim rs01 As New RecordSet
             Set rs01 = rss.Clone
             rs01.MoveFirst
             Do While Not rs01.EOF
                If client <> rs01!B_Clientid Then
                    Exit Sub
                End If
                rs01.movenext
             Loop
        End If
        
    Else
        Exit Sub
    End If
    If TDBGrid2.ApproxCount > 0 Then
        rss.MoveFirst
        Do While Not rss.EOF
            Set rs = New RecordSet
            sql = "select * from G_DraftBillDetailWhite where B_ID='" & id & "' and B_Orderid='" & rss!B_itemid & "'"
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            If rs.RecordCount <= 0 Then
                copytoOne
             End If
             rss.movenext
         Loop
         rss.MoveFirst
    Else
        copytodingdan
        copytodetail
    End If
End Sub

Private Sub copytodingdan()
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    sql = "select * from G_DraftBillWhite where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    sql1 = "delete from G_DraftBillWhite where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql1
End Sub
Private Sub copytodetail()
    Dim rs As New RecordSet
    Dim sql As String
    Set rss1 = New RecordSet
    Dim sql1 As String
    sql = "select * from G_DraftBillDetailWhite "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rss.MoveFirst
    Do While Not rss.EOF
        rs.addnew
        rs!B_datecreate = Now
        rs!B_id = id
        rs!B_GoodsID = rss!B_GoodsNameAlias
        rs!B_Width = rss!B_Width
        rs!B_UnitWeight = rss!B_UnitWeight
        rs!B_qty = rss!B_BreedNum
'        rs!B_price = rss!B_price
        rs!B_PactCode = rss!B_PactCode
        rs!B_ItemIDB = rss!B_ItemIDB
        rs!B_orderid = rss!B_itemid
        rs.Update
        rss.movenext
    Loop
    sql1 = "exec usp_selectWhitedetail_Edit '" & id & "'"
    Debug.Print sql1
    rss1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid2.DataSource = rss1
    setgrid2
End Sub
Private Sub setgrid2()
    TDBGrid2.Columns("B_Name").Caption = "白坯名称"
    TDBGrid2.Columns("B_width").Caption = "门幅"
    TDBGrid2.Columns("B_UnitWeight").Caption = "克重"
    TDBGrid2.Columns("B_Qty").Caption = "数量"
    TDBGrid2.Columns("B_price").Caption = "单价"
    TDBGrid2.Columns("B_sum").Caption = "金额"
    TDBGrid2.Columns("B_DeliveryTime").Caption = "交期"
    TDBGrid2.Columns("B_MemoDetail").Caption = "备注"
    
    TDBGrid2.Columns("B_Name").Locked = True
    TDBGrid2.Columns("B_width").Locked = True
    TDBGrid2.Columns("B_UnitWeight").Locked = True
    TDBGrid2.Columns("B_Qty").Locked = True
    TDBGrid2.Columns("B_sum").Locked = True
    TDBGrid2.Columns("B_DeliveryTime").Locked = True
    

    
    TDBGrid2.Columns("B_pactCode").width = 0
    TDBGrid2.Columns("B_pactCode").Visible = False
    TDBGrid2.Columns("B_pactCode").AllowSizing = False
    TDBGrid2.Columns("B_itemidb").width = 0
    TDBGrid2.Columns("B_itemidb").Visible = False
    TDBGrid2.Columns("B_itemidb").AllowSizing = False
    TDBGrid2.Columns("B_itemid").width = 0
    TDBGrid2.Columns("B_itemid").Visible = False
    TDBGrid2.Columns("B_itemid").AllowSizing = False
    TDBGrid2.Columns("B_ID").width = 0
    TDBGrid2.Columns("B_ID").Visible = False
    TDBGrid2.Columns("B_ID").AllowSizing = False
    TDBGrid2.Columns("B_orderid").width = 0
    TDBGrid2.Columns("B_orderid").Visible = False
    TDBGrid2.Columns("B_orderid").AllowSizing = False
    TDBGrid2.Columns("B_DateCreate").width = 0
    TDBGrid2.Columns("B_DateCreate").Visible = False
    TDBGrid2.Columns("B_DateCreate").AllowSizing = False
    TDBGrid2.Columns("B_SId").width = 0
    TDBGrid2.Columns("B_SId").Visible = False
    TDBGrid2.Columns("B_SId").AllowSizing = False
    TDBGrid2.Columns("B_DeliveryTime").Button = True
    TDBGrid2.MarqueeStyle = dbgHighlightRow
    TDBGrid2.HoldFields
End Sub
'进行修改单价修改金额
Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    Dim a As Double
    a = Val(rss1!B_price) * Val(rss1!B_qty)
   
    Dim sql As String
    sql = "update  G_DraftBillDetailWhite set B_sum='" & a & "',B_price='" & rss1!B_price & "', B_MemoDetail='" & rss1!B_MemoDetail & "' where B_itemid='" & rss1!B_itemid & "'"
    Gm.cnnTool.cnn.Execute sql
    Dim sql2 As String
    Set rss1 = New RecordSet
    sql2 = "exec usp_selectWhitedetail_Edit '" & id & "'"
    rss1.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid2.DataSource = rss1
    setgrid2
    
End Sub

'Private Sub TDBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
''    If ColIndex <> 8 Then
''        Exit Sub
''    End If
''    If KeyAscii < 48 And KeyAscii <> 46 Then
''        Cancel = 1
''        MsgBox "只可以输入数字！"
''    End If
''    If KeyAscii > 57 Then
''        Cancel = 1
''        MsgBox "只可以输入数字！"
''    End If
'
'End Sub


Private Sub TDBGrid2_ButtonClick(ByVal ColIndex As Integer)
    Dim time As String
    Dim a As String
    Dim sql As String
    Dim sql1 As String
    Dim sql2 As String
     If TDBGrid2.Columns("B_DeliveryTime").ColIndex = ColIndex Then
     
        Dim frm1 As New frmpopupTime
        frm1.Show vbModal
        If frm1.bsaved = True Then
            time = Format(frm1.time, "YYYY-MM-DD")
            a = frm1.a
            If a = 1 Then
                sql = "update G_DraftBillDetailWhite set B_DeliveryTime='" & time & "' where B_ID='" & id & "'"
                Gm.cnnTool.cnn.Execute sql
            Else
                sql1 = "update G_DraftBillDetailWhite set B_DeliveryTime='" & time & "' where  B_itemid='" & rss1!B_itemid & "'"
                Debug.Print sql1
                Gm.cnnTool.cnn.Execute sql1
            End If
            Set rss1 = New RecordSet
            sql2 = "exec usp_selectWhitedetail_Edit '" & id & "'"
            rss1.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            TDBGrid2.DataSource = rss1
            setgrid2
        End If
        Unload frm1
    End If
'   rss1.requery
End Sub
Private Sub PushButton2_Click()
    Dim I As Long
    I = 1
    Debug.Print id
    If TDBGrid2.ApproxCount > 0 Then
        rss1.MoveFirst
        Do While Not rss1.EOF
            If IIf(IsNull(rss1!B_price), "", rss1!B_price) = "" Then
                MsgBox "第" & I & "行单价不能为空", vbInformation, "提示"
                Exit Sub
            End If
            If IIf(IsNull(rss1!B_DeliveryTime), "", rss1!B_DeliveryTime) = "" Then
                MsgBox "第" & I & "行交期不能为空", vbInformation, "提示"
                Exit Sub
            End If
            rss1.movenext
            I = I + 1
        Loop
        save
        savedetail
        delete
        FlatEdit1.Text = ""
        FlatEdit3.Text = ""
        client = ""
        grid
    End If
End Sub
Private Sub save()
    Dim sql As String
    Dim f As String
    Debug.Print id
    f = Format(DTPicker1.Value, "YYYY-MM-dd")
    sql = "insert into G_BillWhite (B_ID,B_CodeID,B_BillType,B_Date,B_ContactCom,B_delivery,B_Balance,B_Memo) values('" & id & "','" & GetCodeID & "','WHT10','" & f & "','" & client & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & Trim(FlatEdit4.Text) & "')"
    Gm.cnnTool.cnn.Execute sql
End Sub
Private Sub savedetail()
    Dim sql As String
    Dim sql1 As String
    Dim f As String
    rss1.MoveFirst
    Do While Not rss1.EOF
        sql1 = "update G_WhiteComposition set B_logo='1' where B_itemid='" & rss1!B_orderid & "'"
        Gm.cnnTool.cnn.Execute sql1
        sql = "exec usp_InsertWhiteOrder '" & rss1!B_itemid & "','" & rss1!B_id & "','" & rss1!B_sid & "','" & rss1!B_Width & "','" & rss1!B_UnitWeight & "'"
        sql = sql & ",'" & rss1!B_qty & "','" & rss1!B_price & "','" & rss1!B_sum & "','" & rss1!B_DeliveryTime & "','" & rss1!B_MemoDetail & "','" & rss1!B_PactCode & "','" & rss1!B_orderid & "'"
        Gm.cnnTool.cnn.Execute sql
        rss1.movenext
    Loop
End Sub
Private Sub delete()
    Dim sql As String
    sql = "delete from G_DraftBillDetailWhite where B_id='" & id & "'"
    Gm.cnnTool.cnn.Execute sql
    rss1.requery
End Sub
Private Sub PushButton3_Click()
    Dim I As Long
    I = 1
    Debug.Print id
    If TDBGrid2.ApproxCount > 0 Then
        rss1.MoveFirst
        Do While Not rss1.EOF
            If IIf(IsNull(rss1!B_price), "", rss1!B_price) = "" Then
                MsgBox "第" & I & "行单价不能为空", vbInformation, "提示"
                Exit Sub
            End If
            If IIf(IsNull(rss1!B_DeliveryTime), "", rss1!B_DeliveryTime) = "" Then
                MsgBox "第" & I & "行交期不能为空", vbInformation, "提示"
                Exit Sub
            End If
            rss1.movenext
            I = I + 1
        Loop
        save
        savedetail
        delete
        FlatEdit1.Text = ""
        FlatEdit3.Text = ""
        client = ""
        grid
    Else
        Exit Sub
    End If
    PrintYarn
End Sub

'打印
Private Sub PrintYarn()
    Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmModBLRPreviewOri

    sql = "exec usp_PrintWhiteOrder '" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Set frm1.RecordSet = rs.Clone
    frm1.ObjectID = "22B023"
    frm1.Show
    
End Sub
Private Sub TDBGrid2_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Dim I As Long
    I = 0

    Response = 0
End Sub
'Private Sub JudgeNumeric(ByRef vTDBGrid As TDBGrid, ByVal vColIndex As Long, ByVal vKeyCode As Integer)
'    Dim szTemp As String
'    If vKeyCode < 48 Or vKeyCode > 57 Then
''        MsgBox "只可输入数字"
'        szTemp = Left$(vTDBGrid.Columns(vColIndex).Value, Len(vTDBGrid.Columns(vColIndex).Value) - 1)
'        If Len(szTemp) <= 0 Then
'            vTDBGrid.Columns(vColIndex).Value = 0
'            'vTDBGrid.Columns(vColIndex).
'        Else
'            vTDBGrid.Columns(vColIndex).Value = Val(szTemp)
'            'vTDBGrid.Columns(vColIndex).AllowFocus
'        End If
'
'        vTDBGrid.SetFocus
'    End If
'End Sub
'
'Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
'    JudgeNumeric TDBGrid2, TDBGrid2.Columns("B_price").ColIndex, KeyCode
'End Sub
Private Sub TDBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = TDBGrid2.Columns("B_Price").ColIndex Then
        TDBGrid2.Columns("B_Price").Value = Val(TDBGrid2.Columns("B_Price").Value)
    End If
End Sub
