VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOriginalSelect 
   Caption         =   "构成查询"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13995
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
   ScaleHeight     =   7665
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13995
      _LayoutVersion  =   1
      _ExtentX        =   24686
      _ExtentY        =   13520
      _DataPath       =   ""
      Bands           =   "frmOriginalSelect.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6735
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   12615
         _cx             =   22251
         _cy             =   11880
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
         BorderWidth     =   2
         ChildSpacing    =   2
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
         GridCols        =   5
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmOriginalSelect.frx":2204
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame Frame1 
            Height          =   2160
            Left            =   9465
            TabIndex        =   15
            Top             =   4545
            Width           =   3120
            Begin VB.CommandButton Command3 
               Caption         =   "清空网格"
               Height          =   840
               Left            =   120
               TabIndex        =   18
               Top             =   1200
               Width           =   1350
            End
            Begin VB.CommandButton Command2 
               Caption         =   "删除当前行"
               Height          =   840
               Left            =   1560
               TabIndex        =   17
               Top             =   240
               Width           =   1350
            End
            Begin VB.CommandButton Command1 
               Caption         =   "生成采购订单"
               Height          =   840
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   1350
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1785
            Left            =   30
            ScaleHeight     =   1785
            ScaleWidth      =   12555
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   12555
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   375
               Left            =   8640
               TabIndex        =   13
               Top             =   180
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   4920
               TabIndex        =   3
               Top             =   210
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   231669761
               CurrentDate     =   43058
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   1320
               TabIndex        =   4
               Top             =   210
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   231669761
               CurrentDate     =   43058
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2700
               TabIndex        =   5
               Top             =   1080
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
               Left            =   1320
               TabIndex        =   6
               Top             =   1080
               Width           =   1395
               _Version        =   1048578
               _ExtentX        =   2461
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
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   315
               Left            =   4920
               TabIndex        =   11
               Top             =   1110
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
               Height          =   315
               Left            =   7680
               TabIndex        =   12
               Top             =   210
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "订 单 号:"
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
               Left            =   4020
               TabIndex        =   10
               Top             =   1140
               Width           =   975
               _Version        =   1048578
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "是否执行:"
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   315
               Left            =   420
               TabIndex        =   9
               Top             =   1110
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "供 应 商:"
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
               Left            =   420
               TabIndex        =   8
               Top             =   240
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "起始日期:"
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   195
               Left            =   4020
               TabIndex        =   7
               Top             =   270
               Width           =   795
               _Version        =   1048578
               _ExtentX        =   1402
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "终止日期:"
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
            Height          =   2160
            Left            =   30
            TabIndex        =   14
            Top             =   4545
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3810
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
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   2670
            Left            =   30
            TabIndex        =   19
            Top             =   1845
            Width           =   12555
            _ExtentX        =   22146
            _ExtentY        =   4710
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
            MultiSelect     =   2
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
Attribute VB_Name = "frmOriginalSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rss As RecordSet
Private rsdetail As RecordSet
Private cho As Long

Private supplier As String
Private cls1 As clsGridShow
Public mvarObjectID As String
Private theBLTool As New clsAutoCreateBL
Private Const theObjectID As String = "12B004"
Private Const theObjectID1 As String = "12B006"
Private id As Long


Private Function GetCodeID() As String
    GetCodeID = theBLTool.GetFrameCodeDetail(theObjectID)
End Function
Private Function GetCodeID1() As String
    GetCodeID1 = theBLTool.GetFrameCodeDetail(theObjectID1)
End Function


Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property

'绑定草稿数据
Private Sub setRs()
    Set rsdetail = New RecordSet
    rsdetail.Fields.Append "B_itemid", adVarChar, 100
    rsdetail.Fields.Append "B_itemidb", adVarChar, 100
    rsdetail.Fields.Append "B_suppliers", adVarChar, 100
    rsdetail.Fields.Append "B_goodName", adVarChar, 100
    rsdetail.Fields.Append "B_specification", adVarChar, 100
    rsdetail.Fields.Append "B_qty", adVarChar, 100
    rsdetail.Fields.Append "B_price", adVarChar, 100
    rsdetail.Fields.Append "B_Sum", adVarChar, 100
    rsdetail.Fields.Append "B_client", adVarChar, 100
    rsdetail.Fields.Append "B_memo", adVarChar, 100
    
    rsdetail.Fields.Append "B_suppliersid", adVarChar, 100
    rsdetail.Fields.Append "B_goodid", adVarChar, 100
    rsdetail.Fields.Append "B_clientid", adVarChar, 100
    rsdetail.Fields.Append "B_logo", adVarChar, 100

    rsdetail.Open
    
    TDBGrid2.DataSource = rsdetail
    setrsDetail
End Sub
Private Sub setrsDetail()
    TDBGrid2.Columns("B_itemidb").Caption = "订单号"
    TDBGrid2.Columns("B_suppliers").Caption = "供应商"
    TDBGrid2.Columns("B_goodName").Caption = "原料品名"
    TDBGrid2.Columns("B_specification").Caption = "规格"
    TDBGrid2.Columns("B_qty").Caption = "数量"
    TDBGrid2.Columns("B_price").Caption = "单价"
    TDBGrid2.Columns("B_Sum").Caption = "金额"
    TDBGrid2.Columns("B_client").Caption = "加工户"
    TDBGrid2.Columns("B_memo").Caption = "备注"
    
    TDBGrid2.Columns("B_itemidb").width = 1000
    TDBGrid2.Columns("B_qty").width = 1000
    TDBGrid2.Columns("B_price").width = 1000
    TDBGrid2.Columns("B_Sum").width = 1000
    
       TDBGrid2.Columns("B_itemid").Visible = False
    TDBGrid2.Columns("B_itemid").Locked = True
    TDBGrid2.Columns("B_itemid").AllowSizing = False
    TDBGrid2.Columns("B_suppliersid").Visible = False
    TDBGrid2.Columns("B_suppliersid").Locked = True
    TDBGrid2.Columns("B_suppliersid").AllowSizing = False
        TDBGrid2.Columns("B_goodid").Visible = False
    TDBGrid2.Columns("B_goodid").Locked = True
    TDBGrid2.Columns("B_goodid").AllowSizing = False
        TDBGrid2.Columns("B_clientid").Visible = False
    TDBGrid2.Columns("B_clientid").Locked = True
    TDBGrid2.Columns("B_clientid").AllowSizing = False
           TDBGrid2.Columns("B_logo").Visible = False
    TDBGrid2.Columns("B_logo").Locked = True
    TDBGrid2.Columns("B_logo").AllowSizing = False
    
   TDBGrid2.Columns("B_qty").NumberFormat = "0.0"
   TDBGrid2.Columns("B_price").NumberFormat = "0.00"
   TDBGrid2.Columns("B_Sum").NumberFormat = "0.00"
   TDBGrid2.HoldFields
   TDBGrid2.MarqueeStyle = dbgHighlightRow
End Sub


Private Sub Command1_Click()
    
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    If bol = False Then
        Exit Sub
    End If
    rsdetail.Filter = "B_logo=0"
    If rsdetail.RecordCount > 0 Then
        save
        
        savedetail
    End If
    rsdetail.Filter = ""
    rsdetail.Filter = "B_logo=1"
    If rsdetail.RecordCount > 0 Then
        SaveWhite
        
        Savedetailwhite
    End If
    rsdetail.Filter = ""
    setRs
    MsgBox "生成成功", vbInformation, "提示"
    Grid
End Sub
'打印
Private Sub PrintYarn(ByVal id As String)
    Dim sql As String
    Dim rs As New RecordSet
    
    Dim frm1 As New frmModBLRPreviewOri
    

    sql = "exec usp_PrintYarnOrder '" & id & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

    
    Set frm1.RecordSet = rs.Clone
    frm1.ObjectID = "22B022"
    frm1.Show

End Sub
'打印
Private Sub printwhite(ByVal id As String)
     Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmModBLRPreviewOri

    sql = "exec usp_PrintWhiteOrder '" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Set frm1.RecordSet = rs.Clone
    frm1.ObjectID = "22B023"
    frm1.Show
End Sub

Private Sub Command2_Click()
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    TDBGrid2.delete
    If TDBGrid2.ApproxCount > 0 Then
        TDBGrid2.MoveFirst
    End If
    sumall2
End Sub

Private Sub Command3_Click()
    setRs
    sumall2
End Sub
'验证第二个网格中的供应商是否是一样的
Private Function bol() As Boolean
    bol = True
    Dim rs As New RecordSet
    Set rs = rsdetail.Clone
    Dim a As String
    rs.MoveFirst
    a = rs!B_suppliersid
    Do While Not rs.EOF
        If rs!B_suppliersid <> a Then
            bol = False
            MsgBox "供应商不一样，不能生成", vbInformation, "提示"
            Exit Function
        End If
        rs.movenext
    Loop
    bol = True
End Function

Private Sub Form_Load()
    InitFrm
    zxing
    
    
    '打开窗体就查询
    Grid
    TDBGrid1.SelBookmarks.add 1
    setRs
End Sub

Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    DTPicker2.Value = Now
    supplier = ""
End Sub
Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "查询"
            Grid
        Case "申请打印"
            printdetail
        Case "退出"
            Unload Me
        Case "保存样式"
            setGridStyle
        Case "采购申请"
            setChoose
        Case "审核所有选中"
            setselect
    End Select
End Sub

Private Sub zxing()
    ComboBox1.AddItem "已执行"
    ComboBox1.AddItem "未执行"
    ComboBox1.AddItem "全部"
    ComboBox1.Text = "未执行"
End Sub

Private Sub choose()
    If ComboBox1.Text = "未执行" Then
        cho = 0
    End If
    If ComboBox1.Text = "已执行" Then
        cho = 1
    End If
    If ComboBox1.Text = "全部" Then
        cho = 2
    End If
End Sub

Private Sub Grid()
    Set rss = New RecordSet
    Dim sql As String
    Dim a As String
    Dim b As String
    choose
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    sql = "exec usp_OriginalSelect_Edit '" & a & "','" & b & "','" & supplier & "','" & cho & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "'"
    Debug.Print sql
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid1.DataSource = rss
    setgrid
    sumall
    If rss.RecordCount > 0 Then
'        rss.MoveFirst
'        TDBGrid1.Col = 1
        TDBGrid1.MoveFirst
'        TDBGrid1.SetFocus
         
    End If
End Sub
Private Sub setGridShow()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S009"
        .InitClass TDBGrid1, 3
        .ShowGridFormat
    End With
End Sub
Private Sub setGridStyle()
    Dim i As Long
    Dim strSQL As String
    Dim dWidth As Integer
    Dim szFieldName As String
    
    For i = 0 To TDBGrid1.Columns.Count - 1
        If TDBGrid1.Columns(i).width > 0 Then
            If TDBGrid1.Columns(i).Visible = True Then
                szFieldName = TDBGrid1.Columns(i).DataField
                dWidth = TDBGrid1.Columns(i).width
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S009' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
    
End Sub
Private Sub setgrid()
    setGridShow
'    TDBGrid1.Columns("B_time").Caption = "申请时间"
'    TDBGrid1.Columns("B_itemidb").Caption = "订单号"
'    TDBGrid1.Columns("B_Breed").Caption = "类型"
'    TDBGrid1.Columns("B_storageway").Caption = "单据类型"
'    TDBGrid1.Columns("B_GoodsName").Caption = "品名"
'    TDBGrid1.Columns("B_Width").Caption = "规格"
'    TDBGrid1.Columns("B_clientname").Caption = "供应商"
'    TDBGrid1.Columns("B_kg").Caption = "订单公斤"
'    TDBGrid1.Columns("B_Meter").Caption = "订单米数"
'    TDBGrid1.Columns("B_Qty").Caption = "订单码数"
'    TDBGrid1.Columns("B_KG01").Caption = "申请公斤"
'    TDBGrid1.Columns("B_KG02").Caption = "白坯计划公斤"
'    TDBGrid1.Columns("B_logo").Caption = "是否执行"
'    TDBGrid1.Columns("B_transfersitemidb").Caption = "调拨单号"
    TDBGrid1.Columns("SDate").Visible = False
    TDBGrid1.Columns("SDate").Locked = True
    TDBGrid1.Columns("SDate").AllowSizing = False
     TDBGrid1.Columns("Edate").Visible = False
    TDBGrid1.Columns("Edate").Locked = True
    TDBGrid1.Columns("Edate").AllowSizing = False
     TDBGrid1.Columns("ClientName").Visible = False
    TDBGrid1.Columns("ClientName").Locked = True
    TDBGrid1.Columns("ClientName").AllowSizing = False
     TDBGrid1.Columns("dingdan").Visible = False
    TDBGrid1.Columns("dingdan").Locked = True
    TDBGrid1.Columns("dingdan").AllowSizing = False
       TDBGrid1.Columns("B_itemid").Visible = False
    TDBGrid1.Columns("B_itemid").Locked = True
    TDBGrid1.Columns("B_itemid").AllowSizing = False
    TDBGrid1.Columns("B_SID").Visible = False
    TDBGrid1.Columns("B_SID").Locked = True
    TDBGrid1.Columns("B_SID").AllowSizing = False
    TDBGrid1.Columns("B_Specification").Visible = False
    TDBGrid1.Columns("B_Specification").Locked = True
    TDBGrid1.Columns("B_Specification").AllowSizing = False
    TDBGrid1.Columns("B_pactcode").Visible = False
    TDBGrid1.Columns("B_pactcode").Locked = True
    TDBGrid1.Columns("B_pactcode").AllowSizing = False
    TDBGrid1.Columns("B_suppliers").Visible = False
    TDBGrid1.Columns("B_suppliers").Locked = True
    TDBGrid1.Columns("B_suppliers").AllowSizing = False
    
    TDBGrid1.Columns("B_Goodsid").Visible = False
    TDBGrid1.Columns("B_Goodsid").Locked = True
    TDBGrid1.Columns("B_Goodsid").AllowSizing = False
    TDBGrid1.Columns("B_clientid").Visible = False
    TDBGrid1.Columns("B_clientid").Locked = True
    TDBGrid1.Columns("B_clientid").AllowSizing = False
    
    
    
    
    'TDBGridMergeCell TDBGrid1, "B_itemidb"
    TDBGridMergeCell TDBGrid1, "B_itemidbHidden"
    
    TDBGridMergeCell TDBGrid1, "B_kg"
    TDBGridMergeCell TDBGrid1, "B_Meter"
    TDBGridMergeCell TDBGrid1, "B_Qty"
    TDBGridMergeCell TDBGrid1, "B_KG01"
    TDBGridMergeCell TDBGrid1, "B_KG02"
    
    
    TDBGrid1.Columns("B_logo").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.HoldFields
    TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub PushButton1_Click()
        Dim frm1 As New frmpopupSuppliers
'        frm1.ContactType = "原料供应商"
        frm1.Show vbModal
        supplier = frm1.clientid
        FlatEdit1.Text = frm1.ClientName
        Unload frm1
End Sub

Private Sub printdetail()
    On Error Resume Next
    
    Dim rs As New RecordSet
    Dim sql As String
        Dim a As String
    Dim b As String
    choose
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    sql = "exec usp_OriginalSelect_Edit '" & a & "','" & b & "','" & supplier & "','" & cho & "','" & FlatEdit1.Text & "','" & FlatEdit2.Text & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Debug.Print sql
    Dim frm1 As New frmModBLRPreviewOri
    
    Set frm1.RecordSet = rs.Clone
    frm1.ObjectID = "22b029"
    frm1.Show
    
'    Unload frm1
End Sub

Private Sub sumall()
     Dim rs As New RecordSet
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As Double
    Dim f As String
    Dim g As String
    Dim h As String
    Dim i As String
    Dim j As String
    Dim k As Double
    Dim m As String
    If rss.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    e = 0
    k = 0
    Set rs = rss.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_kg), 0, rs!B_kg)
        b = b + IIf(IsNull(rs!B_meter), 0, rs!B_meter)
        c = c + IIf(IsNull(rs!B_qty), 0, rs!B_qty)
        d = d + IIf(IsNull(rs!B_KG01), 0, rs!B_KG01)
        e = e + IIf(IsNull(rs!B_KG02), 0, rs!B_KG02)
        k = k + IIf(IsNull(rs!B_BreedNum), 0, rs!B_BreedNum)
        rs.movenext
    Loop
    f = Format(a, "0.0")
    g = Format(b, "0.0")
    h = Format(c, "0.0")
    i = Format(d, "0.0")
    j = Format(e, "0.0")
    m = Format(k, "0.0")
    TDBGrid1.Columns("B_itemidb").FooterText = "合计"
    TDBGrid1.Columns("B_kg").FooterText = "" & f & ""
    TDBGrid1.Columns("B_Meter").FooterText = "" & g & ""
    TDBGrid1.Columns("B_qty").FooterText = "" & h & ""
    TDBGrid1.Columns("B_KG01").FooterText = "" & i & ""
    TDBGrid1.Columns("B_KG02").FooterText = "" & j & ""
    TDBGrid1.Columns("B_BreedNum").FooterText = "" & m & ""
    rs.MoveFirst
End Sub

'设置弹出保存数据
Private Sub setChoose()
    On Error GoTo IFERR
    
    
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    If rss!B_ItemID = "" Then
        Exit Sub
    End If
    
    Dim frm1 As New frmOriginalSelect_Edit
    frm1.FlatEdit3.Text = rss!B_BreedNum
    frm1.FlatEdit4.Text = Val(rss!B_BreedNum) - fp(rss!B_ItemID, rss!B_Breed)
    frm1.FlatEdit5.Text = Val(rss!B_BreedNum) - fp(rss!B_ItemID, rss!B_Breed)
    frm1.Show vbModal
    If frm1.bool = False Then
        Exit Sub
    End If
    rsdetail.AddNew
        rsdetail!B_ItemID = rss!B_ItemID
        rsdetail!B_ItemIDB = rss!B_ItemIDB
        rsdetail!B_suppliers = rss!B_ClientName
        rsdetail!B_goodName = rss!B_GoodsName
        rsdetail!B_specification = rss!B_Width
        rsdetail!B_qty = frm1.FlatEdit5.Text
        rsdetail!B_price = frm1.FlatEdit1.Text
        rsdetail!B_sum = Val(frm1.FlatEdit1.Text) * Val(frm1.FlatEdit5.Text)
        rsdetail!B_Client = frm1.FlatEdit2.Text
        rsdetail!B_memo = frm1.FlatEdit6.Text
        
        rsdetail!B_suppliersid = rss!B_Clientid
        rsdetail!B_goodid = rss!B_GoodsID
        rsdetail!B_Clientid = frm1.Originalsuppliers
        If rss!B_Breed = "白坯" Then
            rsdetail!B_logo = 1
        Else
            rsdetail!B_logo = 0
        End If
    rsdetail.Update
    sumall2
    Unload frm1
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "选择一行数据" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
    
End Sub
Private Sub setselect()
    On Error GoTo IFERR
    Dim tdbgRow As Variant
    Dim i As Long
    Dim a As String
    i = 1
    For Each tdbgRow In TDBGrid1.SelBookmarks
        rss.bookmark = tdbgRow
       If i = 1 Then
            a = rss!B_Clientid
       End If
       If i > 1 Then
        If a <> rss!B_Clientid Then
            Exit Sub
        End If
       End If
       i = i + 1
    Next
    
    For Each tdbgRow In TDBGrid1.SelBookmarks
        rss.bookmark = tdbgRow
        
        rsdetail.AddNew
        rsdetail!B_ItemID = rss!B_ItemID
        rsdetail!B_ItemIDB = rss!B_ItemIDB
        rsdetail!B_suppliers = rss!B_ClientName
        rsdetail!B_goodName = rss!B_GoodsName
        rsdetail!B_specification = rss!B_Width
        rsdetail!B_qty = rss!B_BreedNum
        rsdetail!B_price = ""
        rsdetail!B_sum = ""
        rsdetail!B_Client = ""
        rsdetail!B_memo = ""
        
        rsdetail!B_suppliersid = rss!B_Clientid
        rsdetail!B_goodid = rss!B_GoodsID
        rsdetail!B_Clientid = ""
        If rss!B_Breed = "白坯" Then
            rsdetail!B_logo = 1
        Else
            rsdetail!B_logo = 0
        End If
        rsdetail.Update
    Next
     Exit Sub
IFERR:
    Dim szErr As String
    szErr = "选择一行数据" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'获取采购申请弹出页面的已分配数量
Private Function fp(ByVal a As String, ByVal c As String) As Long
    Dim rs1 As New RecordSet
    Set rs1 = rsdetail.Clone
    Dim sql As String
    Dim rs As New RecordSet
    
    Dim b As Long
    If c = "原料" Then
        sql = "SELECT  isnull(SUM(B_Qty),0) AS B_sum FROM G_BillDetailYarn a LEFT OUTER JOIN G_BillYarn b"
        sql = sql & " ON a.B_ID=b.B_ID WHERE B_orderid='" & a & "' AND B_BillType='YARN08' AND B_ObjectID='12B004'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Else
        sql = "SELECT  isnull(SUM(B_kg),0) AS B_sum FROM G_BillDetailwhite a LEFT OUTER JOIN G_Billwhite b"
        sql = sql & " ON a.B_ID=b.B_ID WHERE B_orderid='" & a & "' AND B_BillType='WHT10'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    End If
    Debug.Print sql
    b = rs!B_sum
    If TDBGrid2.ApproxCount > 0 Then
        rs1.MoveFirst
        
        Do While Not rs1.EOF
            If rs1!B_ItemID = a Then
                 b = b + Val(rs1!B_qty)
            End If
            rs1.movenext
        Loop
    End If
    fp = b
End Function

Private Sub save()
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    
    Dim sql2 As String
    Dim f As String
    Dim a As String
    
     sql = "select * from G_DraftBillYarn where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    sql2 = "delete from G_DraftBillYarn where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql2
    
    f = Now
    rsdetail.MoveFirst
    a = rsdetail!B_suppliersid
    
    sql1 = "insert into G_BillYarn (B_ID,B_CodeID,B_BillType,B_Date,B_ObjectID,B_ContactCom) "
    sql1 = sql1 & "values('" & id & "','" & GetCodeID & "','YARN08','" & f & "','12B004','" & a & "')"
    Debug.Print sql1
    Gm.cnnTool.cnn.Execute sql1
End Sub
Private Sub savedetail()
   Dim sql As String
   Dim rs As New RecordSet
   Dim sql1 As String
   Dim itemid As String
   Dim sql3 As String
   sql = "select * from G_DraftBillDetailYarn where 1=1"
   rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
   
   rsdetail.MoveFirst
   Do While Not rsdetail.EOF
        If rsdetail!B_logo = 0 Then
            sql1 = "update G_WhiteComposition set B_logo='1' where B_itemid='" & rss!B_ItemID & "'"
            Gm.cnnTool.cnn.Execute sql1
            rs.AddNew
            rs!B_datecreate = Now
            rs.Update
            itemid = rs!B_ItemID
            
            sql3 = "exec usp_InsertOriginal_SQ '" & itemid & "','" & id & "','" & rsdetail!B_Clientid & "','" & rsdetail!B_goodid & "','" & rsdetail!B_specification & "'"
            sql3 = sql3 & ",'" & rsdetail!B_qty & "','" & rsdetail!B_price & "','" & rsdetail!B_sum & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_memo & "','" & rsdetail!B_ItemID & "'"
            Gm.cnnTool.cnn.Execute sql3
            
            
             sql4 = "delete from G_DraftBillDetailYarn where B_itemid='" & itemid & "'"
            Gm.cnnTool.cnn.Execute sql4
        End If
        rsdetail.movenext
   Loop
   PrintYarn (id)
End Sub
Private Sub sumall2()
    Dim rs As New RecordSet
    Dim a As Double
    Dim b As Double
    Dim c As String
    Dim d As String
   
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
 
    Set rs = rsdetail.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_qty), 0, rs!B_qty)
        b = b + IIf(IsNull(Val(rs!B_sum)), 0, Val(rs!B_sum))
        
        
        rs.movenext
    Loop
    c = Format(a, "0.0")
    d = Format(b, "0.00")

    TDBGrid2.Columns("B_ItemIDB").FooterText = "合计"
    
    TDBGrid2.Columns("B_qty").FooterText = "" & c & ""
    TDBGrid2.Columns("B_sum").FooterText = "" & d & ""
    rs.MoveFirst
End Sub
Private Sub TDBGridMergeCell(ByRef vTDBGrid As TDBGrid, ByVal vFieldName As String)

'    vTDBGrid.Columns(vFieldName).Merge = dbgMergeFree
    vTDBGrid.Columns(vFieldName).Merge = dbgMergeRestricted
    

End Sub

Private Sub SaveWhite()
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    
    Dim sql2 As String
    Dim f As String
    Dim a As String
    
     sql = "select * from G_DraftBillwhite where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_id
    sql2 = "delete from G_DraftBillwhite where B_ID='" & id & "'"
    Gm.cnnTool.cnn.Execute sql2
    
    f = Now
    rsdetail.MoveFirst
    a = rsdetail!B_suppliersid
    
    sql1 = "insert into G_Billwhite (B_ID,B_CodeID,B_BillType,B_Date,B_ObjectID,B_ContactCom) "
    sql1 = sql1 & "values('" & id & "','" & GetCodeID1 & "','WHT10','" & f & "','12B006','" & a & "')"
    Debug.Print sql1
    Gm.cnnTool.cnn.Execute sql1
End Sub
Private Sub Savedetailwhite()
   Dim sql As String
   Dim rs As New RecordSet
   Dim sql1 As String
   Dim itemid As String
   Dim sql3 As String
   sql = "select * from G_DraftBillDetailwhite where 1=1"
   rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
   
   rsdetail.MoveFirst
   Do While Not rsdetail.EOF
        If rsdetail!B_logo = 1 Then
            sql1 = "update G_WhiteComposition set B_logo='1' where B_itemid='" & rss!B_ItemID & "'"
            Gm.cnnTool.cnn.Execute sql1
            rs.AddNew
            rs!B_datecreate = Now
            rs.Update
            itemid = rs!B_ItemID
            
            sql3 = "exec usp_Insertwhite_SQ '" & itemid & "','" & id & "','" & rsdetail!B_Clientid & "','" & rsdetail!B_goodid & "','" & rsdetail!B_specification & "'"
            sql3 = sql3 & ",'" & rsdetail!B_qty & "','" & rsdetail!B_price & "','" & rsdetail!B_sum & "','" & rsdetail!B_ItemIDB & "','" & rsdetail!B_memo & "','" & rsdetail!B_ItemID & "'"
            Gm.cnnTool.cnn.Execute sql3
            
            
             sql4 = "delete from G_DraftBillDetailwhite where B_itemid='" & itemid & "'"
            Gm.cnnTool.cnn.Execute sql4
        End If
        rsdetail.movenext
   Loop
   printwhite (id)
End Sub
