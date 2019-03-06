VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmColorfactoryReport 
   Caption         =   "家纺厂序时表"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11640
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
   ScaleHeight     =   7590
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10725
      _LayoutVersion  =   1
      _ExtentX        =   18918
      _ExtentY        =   12965
      _DataPath       =   ""
      Bands           =   "frmColorfactoryReport.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6495
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   9915
         _cx             =   17489
         _cy             =   11456
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
         GridRows        =   4
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmColorfactoryReport.frx":2906
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1440
            Left            =   30
            ScaleHeight     =   1440
            ScaleWidth      =   9855
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   9855
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   4800
               TabIndex        =   3
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Format          =   232062977
               CurrentDate     =   43110
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   1560
               TabIndex        =   4
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Format          =   232062977
               CurrentDate     =   43110
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   315
               Left            =   4800
               TabIndex        =   5
               Top             =   990
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
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2880
               TabIndex        =   6
               Top             =   960
               Width           =   375
               _Version        =   1048578
               _ExtentX        =   661
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               BackColor       =   16777215
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   1560
               TabIndex        =   7
               Top             =   960
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
            Begin XtremeSuiteControls.ComboBox ComboBox2 
               Height          =   390
               Left            =   8100
               TabIndex        =   13
               Top             =   240
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
            Begin XtremeSuiteControls.Label Label16 
               Height          =   375
               Left            =   6600
               TabIndex        =   14
               Top             =   255
               Width           =   1455
               _Version        =   1048578
               _ExtentX        =   2566
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "查询日期样式:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   3840
               TabIndex        =   11
               Top             =   1020
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "单据类型:"
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   480
               TabIndex        =   10
               Top             =   300
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "起始日期:"
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   3840
               TabIndex        =   9
               Top             =   300
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "终止日期:"
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Left            =   480
               TabIndex        =   8
               Top             =   1020
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "家 纺 厂:"
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   4965
            Left            =   30
            TabIndex        =   12
            Top             =   1500
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   8758
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
            Splits(0).FilterBar=   -1  'True
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
      End
   End
End
Attribute VB_Name = "frmColorfactoryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rss As RecordSet
Public mvarObjectID As String
Private Originalsuppliers As String '供应商的id
Private ch As String
Private m_DateFormat As Long
Private prtRS As New RecordSet
Private GetString As String             '获取一个全局筛选的字符串


Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "打印"
            Printstyle
        Case "打印筛选"
            Printstyle1
        Case "查询"
            Grid
        Case "退出"
            Unload Me
        Case "保存样式"
            setGridStyle
        Case "入库"
            colorinsert
        Case "发货"
            colorreturn
            
    End Select
End Sub

Private Sub ComboBox2_Click()
    If ComboBox2.Text = "账目日期" Then
        m_DateFormat = 1
    End If
    If ComboBox2.Text = "制单日期" Then
        m_DateFormat = 0
    End If
   Grid
End Sub
Private Sub Valuation()
    ComboBox2.AddItem "账目日期"
    ComboBox2.AddItem "制单日期"
    ComboBox2.Text = "制单日期"
    m_DateFormat = 0
End Sub

Private Sub Form_Load()
    InitFrm
    cob1
    Valuation
    '打开窗体就执行查询
    Grid
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
   
     DTPicker1.Value = DateSerial(Year(Now), Month(Now), 1)
     DTPicker2.Value = Now
End Sub
Private Sub PushButton1_Click()
    Dim frm1 As New frmPopupDanWei
    frm1.ContactType = "家纺厂"
    frm1.Show vbModal
    Originalsuppliers = frm1.clientid
    FlatEdit3.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub cob1()
    ComboBox1.Clear
    ComboBox1.AddItem "色布发家纺厂"
    ComboBox1.AddItem "家纺厂成品入库"
    ComboBox1.AddItem "全部"
    ComboBox1.Text = "全部"
End Sub
Private Sub chose()
    If ComboBox1.Text = "全部" Then
        ch = 2
    End If
    If ComboBox1.Text = "色布发家纺厂" Then
        ch = "COL18"
    End If
    If ComboBox1.Text = "家纺厂成品入库" Then
        ch = "COL19"
    End If
End Sub

Private Sub Grid()
    Set rss = New RecordSet
    Dim sql As String
    Dim a As String
    Dim b As String
    chose
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    sql = "exec usp_colorfactoryReport '" & a & "','" & b & "','" & Originalsuppliers & "','" & ch & "','" & Gm.SysID.SystemUserName & "','" & m_DateFormat & "'"
  Debug.Print sql
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

    TDBGrid1.DataSource = rss
    
    SetGrid
End Sub
Private Sub SetGrid()
    setGridShow
'    TDBGrid1.Columns("B_PS").NumberFormat = "0.0"
    TDBGrid1.Columns("B_KG").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_Sum").NumberFormat = "0.00"
    TDBGrid1.HoldFields
'    TDBGrid1.MarqueeStyle = dbgHighlightRow
    TDBGrid1.MarqueeStyle = dbgFloatingEditor

    TDBGrid1.FetchRowStyle = True
    TDBGrid1.Columns("B_Prepaidfreight").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_Waitfreight").ValueItems.Presentation = dbgCheckBox
    
    TDBGrid1.Columns("username").Visible = False
    TDBGrid1.Columns("username").Locked = True
    TDBGrid1.Columns("username").AllowSizing = False
       TDBGrid1.Columns("Sdate").Visible = False
    TDBGrid1.Columns("Sdate").Locked = True
    TDBGrid1.Columns("Sdate").AllowSizing = False
           TDBGrid1.Columns("edate").Visible = False
    TDBGrid1.Columns("edate").Locked = True
    TDBGrid1.Columns("edate").AllowSizing = False
    
    TDBGridMergeCell TDBGrid1, "B_XHBL"
    TDBGridMergeCell TDBGrid1, "B_CodeID"
    TDBGrid1.Columns("B_Hex").FetchStyle = True
    sumall
End Sub
Private Sub setGridShow()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S065"
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
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S065' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
End Sub

'打印
Private Sub Printstyle()
        If TDBGrid1.ApproxCount <= 0 Then
            Exit Sub
        End If

        Dim rs As New RecordSet
        Dim sql As String
        Dim a As String
        Dim b As String
        chose
        a = Format(DTPicker1.Value, "YYYY-MM-DD")
        b = Format(DTPicker2.Value, "YYYY-MM-DD")
        sql = "exec usp_colorfactoryReport '" & a & "','" & b & "','" & Originalsuppliers & "','" & ch & "','" & Gm.SysID.SystemUserName & "','" & m_DateFormat & "'"
        
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Dim frm1 As New frmModBLRPreviewOri
        Set frm1.RecordSet = rs.Clone
        frm1.obj = "11S065"
        frm1.ObjectID = "22B087"
        frm1.Show
End Sub
  '打印筛选
Private Sub Printstyle1()
        If TDBGrid1.ApproxCount <= 0 Then
            Exit Sub
        End If

        Dim rs As New RecordSet
        Dim sql As String
        Dim a As String
        Dim b As String
        chose
        a = Format(DTPicker1.Value, "YYYY-MM-DD")
        b = Format(DTPicker2.Value, "YYYY-MM-DD")
        sql = "exec usp_colorfactoryReport '" & a & "','" & b & "','" & Originalsuppliers & "','" & ch & "','" & Gm.SysID.SystemUserName & "','" & m_DateFormat & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Dim cls1 As New clsRecordset
       Set prtRS = cls1.buildRsWithData(rs, GetString).Clone
        
        Dim frm1 As New frmModBLRPreviewOri
        Set frm1.RecordSet = prtRS.Clone
        frm1.obj = "11S065"
        frm1.ObjectID = "22B087"
        frm1.Show
End Sub
'构建过滤字符串
Public Function buildFilterString() As String
    Dim i As Long
    Dim j As Long
    Dim szReturn As String
    
    szReturn = ""
    j = 0
    For i = 0 To TDBGrid1.Columns.Count - 1
        If Len(Trim(TDBGrid1.Columns(i).FilterText)) > 0 Then
            If j > 0 Then
                szReturn = szReturn & " And "
            End If
        
            szReturn = szReturn & TDBGrid1.Columns(i).DataField & " like '%" & TDBGrid1.Columns(i).FilterText & "%'"
            j = j + 1
        End If
    
    Next
    
    Debug.Print szReturn
    buildFilterString = szReturn
End Function


   Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, _
    ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    
    Dim lState As String
    lState = IIf(IsNull(TDBGrid1.Columns("B_BillName").CellValue(bookmark)), "", TDBGrid1.Columns("B_BillName").CellValue(bookmark))
    
    
    If lState = "色布发家纺厂" Then
        RowStyle.BackColor = &HC0FFFF
    End If
    
End Sub
Private Sub sumall()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim a As Double
    Dim b As Double
    Dim c As String
    Dim d As String
    a = 0
    b = 0
    TDBGrid1.Columns("B_XHBL").FooterText = "合计"
    If rss.RecordCount <= 0 Then
        TDBGrid1.Columns("B_ps").FooterText = "" & a & ""
        TDBGrid1.Columns("B_KG").FooterText = "" & b & ""
    End If
    
    
    Set rs = rss.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_ps), 0, rs!B_ps)
        b = b + IIf(IsNull(rs!B_kg), 0, rs!B_kg)
        rs.movenext
    Loop
    c = Format(a, "0.0")
'    d = Format(b, "0.00")
    TDBGrid1.Columns("B_ps").FooterText = "" & c & ""
    TDBGrid1.Columns("B_kg").FooterText = "" & b & ""
End Sub
Private Sub sumall1()
    Dim Sum As Double
    Dim Sum2 As Double
    
    Dim i As Integer
    Dim n As Long
    n = TDBGrid1.ApproxCount
    Sum = 0
    Sum2 = 0
   
    For i = 1 To n
        Sum = Sum + Val(TDBGrid1.Columns("B_ps").Text)
        Sum2 = Sum2 + Val(TDBGrid1.Columns("B_kg").Text)
        
        TDBGrid1.movenext
    Next i
    TDBGrid1.Columns("B_ps").FooterText = "" & Sum & ""
    TDBGrid1.Columns("B_kg").FooterText = "" & Sum2 & ""
    
End Sub
Private Sub TDBGridMergeCell(ByRef vTDBGrid As TDBGrid, ByVal vFieldName As String)

    vTDBGrid.Columns(vFieldName).Merge = dbgMergeFree
 

End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

   
    On Error Resume Next
    Debug.Print TDBGrid1.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid1.Columns("B_Hex").CellValue(bookmark)
End Sub

'双击打开打开原单
Private Sub TDBGrid1_DblClick()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
 

    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_BillColor where B_id='" & rss!B_id & "'"
 
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Debug.Print rs!B_BillType
    If Trim(rs!B_BillType) = "COL18" Then '色布发家纺厂
        openbill rss!B_id
    End If
    If Trim(rs!B_BillType) = "COL19" Then '家纺厂成品入库
        openbill1 rss!B_id
    End If

    
End Sub
'打开页面显示上面内容,,,,,色布入库
Private Sub openbill(ByVal id As String)
    Dim a As Long
    a = TDBGrid1.bookmark
     Dim rs As New RecordSet
    Dim sql As String
   sql = "select  B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope,B_Freighttelephone,B_drivename,B_CostPay"
    sql = sql & " from G_Billcolor a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_id='" & id & "'"

    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'    If rs.RecordCount <= 0 Then
'        MsgBox "无订单信息", vbInformation + vbOKOnly, "提示"
'        Exit Sub
'    End If
    Dim frm1 As New frmColorOutfactory
    
     frm1.id = rs!B_id
    frm1.FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    frm1.DTPicker1.Value = rs!B_Date
    frm1.FlatEdit2.Text = IIf(IsNull(rs!B_Freight), "", rs!B_Freight)
    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_PlateNumber), "", rs!B_PlateNumber)
    frm1.FlatEdit4.Text = IIf(IsNull(rs!B_PNumber), "", rs!B_PNumber)
    frm1.ComboBox1.Text = rs!B_payment
    frm1.FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    frm1.Originalsuppliers = IIf(IsNull(rs!B_Shipment), "", rs!B_Shipment)
    frm1.FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    frm1.fh = rs!B_Hand
     frm1.ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
     frm1.FlatEdit7.Text = IIf(IsNull(rs!B_Freighttelephone), "", rs!B_Freighttelephone)
    frm1.FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    If rs!B_costpay = 0 Then
        frm1.ComboBox3.Text = "否"
    End If
    If rs!B_costpay = 1 Then
        frm1.ComboBox3.Text = "是"
    End If
    If rs!B_costpay = "" Then
        frm1.ComboBox3.Text = ""
    End If
    frm1.openbill
    frm1.Show vbModal
    Grid
    
    If rss.RecordCount > 0 Then
        If rss.RecordCount < a Then
            rss.movelast
        Else
            TDBGrid1.bookmark = a
        End If
    End If
End Sub
'打开页面显示上面内容,,,,,色布退货
Private Sub openbill1(ByVal id As String)
    Dim a As Long
    a = TDBGrid1.bookmark
    Dim rs As New RecordSet
    Dim sql As String
   sql = "select  B_id,B_CodeID,B_Date,B_Freight,B_PlateNumber,B_PNumber,B_payment,B_Shipment,b.B_ClientName,B_Hand,c.B_Name,B_cope,B_Freighttelephone,B_drivename,B_CostPay"
    sql = sql & " from G_Billcolor a left outer join G_ContactCompany b on a.B_Shipment=b.B_ClientID"
    sql = sql & " left outer join G_Employee c on a.B_Hand=c.B_SID where B_id='" & id & "'"

    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'    If rs.RecordCount <= 0 Then
'        MsgBox "无订单信息", vbInformation + vbOKOnly, "提示"
'        Exit Sub
'    End If
    Dim frm1 As New frmColorfactory
    
     frm1.id = rs!B_id
    frm1.FlatEdit3.Text = IIf(IsNull(rs!B_Codeid), "", rs!B_Codeid)
    frm1.DTPicker1.Value = rs!B_Date
    frm1.FlatEdit2.Text = IIf(IsNull(rs!B_Freight), "", rs!B_Freight)
    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_PlateNumber), "", rs!B_PlateNumber)
    frm1.FlatEdit4.Text = IIf(IsNull(rs!B_PNumber), "", rs!B_PNumber)
    frm1.ComboBox1.Text = rs!B_payment
    frm1.FlatEdit1.Text = IIf(IsNull(rs!B_ClientName), "", rs!B_ClientName)
    frm1.Originalsuppliers = IIf(IsNull(rs!B_Shipment), "", rs!B_Shipment)
    frm1.FlatEdit5.Text = IIf(IsNull(rs!B_name), "", rs!B_name)
    frm1.fh = rs!B_Hand
     frm1.ComboBox2.Text = IIf(IsNull(rs!B_cope), "", rs!B_cope)
    frm1.FlatEdit7.Text = IIf(IsNull(rs!B_Freighttelephone), "", rs!B_Freighttelephone)
    frm1.FlatEdit8.Text = IIf(IsNull(rs!B_drivename), "", rs!B_drivename)
    If rs!B_costpay = 0 Then
        frm1.ComboBox3.Text = "否"
    End If
    If rs!B_costpay = 1 Then
        frm1.ComboBox3.Text = "是"
    End If
    If rs!B_costpay = "" Then
        frm1.ComboBox3.Text = ""
    End If
    frm1.openbill
    frm1.Show vbModal
    Grid
    
    If rss.RecordCount > 0 Then
        If rss.RecordCount < a Then
            rss.movelast
        Else
            TDBGrid1.bookmark = a
        End If
    End If
End Sub


Private Sub TDBGrid1_FilterChange()
'    Dim oldCol As Long
'    oldCol = TDBGrid1.col
'    ExeTDBGridFilterChange TDBGrid1, rs
'    TDBGrid1.col = oldCol

    
    ExeTDBGridFilterChange TDBGrid1, rss
    Dim lColIndex As Long
    lColIndex = TDBGrid1.Col
    sumall1
        TDBGrid1.SetFocus
    TDBGrid1.Col = lColIndex
    TDBGrid1.FilterActive = True
    TDBGrid1.SelStart = Len(TDBGrid1.Columns(lColIndex).FilterText)
End Sub


Private Sub ExeTDBGridFilterChange(ByRef vTDBGrid As TDBGrid, ByRef vRs As RecordSet)
    On Error GoTo IFERR
    Dim Col As Integer
    Col = vTDBGrid.Col
       
    vTDBGrid.HoldFields
    vRs.Filter = GetTDBGridFilterString(vTDBGrid)
    vTDBGrid.Col = Col
    vTDBGrid.EditActive = True
       
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "错误发生于对网格控件进行过滤中" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
  
Private Function GetTDBGridFilterString(ByRef vTDBGrid As TDBGrid) As String
    On Error Resume Next
    Dim tmp As String
    Dim n As Integer
    Dim Col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
       
    Set cols = vTDBGrid.Columns
       
    For Each Col In cols
        If Trim(Col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            Select Case Col.DataWidth
                Case 23, 6, 11
                    tmp = tmp & Col.DataField & " =" & Col.FilterText
                Case Else
                    tmp = tmp & Col.DataField & " LIKE '%" & Col.FilterText & "%'"
            End Select
        End If
    Next Col
                      
     GetString = tmp
    GetTDBGridFilterString = tmp
End Function


Private Sub colorinsert()
    
    Dim frm1 As New frmColorfactory
    frm1.Show vbModal
    If frm1.bol = True Then
        Grid
        TDBGrid1.MoveFirst
    End If
End Sub

Private Sub colorreturn()
    Dim frm1 As New frmColorOutfactory
    frm1.Show vbModal
    If frm1.bol = True Then
        Grid
        TDBGrid1.MoveFirst
    End If
End Sub

'设置网格控件的头部和脚部可以点击
Private Sub SetGridHeadButton()
    Dim obcol As TrueOleDBGrid80.Column
    For Each obcol In TDBGrid1.Columns
        'obcol.ButtonFooter = True
        obcol.ButtonHeader = True
    Next obcol
End Sub
  
Private Sub TDBGrid1_HeadClick(ByVal colIndex As Integer)
    Static m_bSortFlag As Boolean
       
    If m_bSortFlag = False Then
    '顺序排列
        rss.Sort = TDBGrid1.Columns(colIndex).DataField
        m_bSortFlag = True
    Else
    '逆序排列
        rss.Sort = TDBGrid1.Columns(colIndex).DataField & " DESC"
        m_bSortFlag = False
    End If
End Sub

