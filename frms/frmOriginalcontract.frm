VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOriginalcontract 
   Caption         =   " "
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12930
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
   ScaleHeight     =   8280
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   8280
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12930
      _LayoutVersion  =   1
      _ExtentX        =   22807
      _ExtentY        =   14605
      _DataPath       =   ""
      Bands           =   "frmOriginalcontract.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7530
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   660
         Width           =   12900
         _cx             =   22754
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
         GridCols        =   5
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmOriginalcontract.frx":2C78
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1785
            Left            =   10695
            ScaleHeight     =   1785
            ScaleWidth      =   2160
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   45
            Width           =   2160
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   495
               Left            =   540
               TabIndex        =   21
               Top             =   1260
               Width           =   1335
               _Version        =   1048578
               _ExtentX        =   2355
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "生成运费"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   720
               TabIndex        =   19
               Top             =   120
               Width           =   1275
               _Version        =   1048578
               _ExtentX        =   2249
               _ExtentY        =   661
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
               BackColor       =   16777215
            End
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   315
               Left            =   720
               TabIndex        =   22
               Top             =   780
               Width           =   1275
               _Version        =   1048578
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   16777215
               Style           =   2
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   840
               Width           =   435
               _Version        =   1048578
               _ExtentX        =   767
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "运费:"
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   180
               Width           =   435
               _Version        =   1048578
               _ExtentX        =   767
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "车号:"
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1785
            Left            =   45
            ScaleHeight     =   1785
            ScaleWidth      =   8250
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   8250
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2520
               TabIndex        =   15
               Top             =   570
               Width           =   435
               _Version        =   1048578
               _ExtentX        =   767
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   ".."
               UseVisualStyle  =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   4980
               TabIndex        =   3
               Top             =   600
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               Format          =   199753729
               CurrentDate     =   43059
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   1260
               TabIndex        =   4
               Top             =   570
               Width           =   1275
               _Version        =   1048578
               _ExtentX        =   2249
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
               TabIndex        =   5
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
               Left            =   4980
               TabIndex        =   6
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
               Left            =   8340
               TabIndex        =   7
               Top             =   600
               Width           =   2175
               _Version        =   1048578
               _ExtentX        =   3836
               _ExtentY        =   1402
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               MultiLine       =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   435
               Left            =   4020
               TabIndex        =   13
               Top             =   120
               Width           =   2595
               _Version        =   1048578
               _ExtentX        =   4577
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "凯鑫原料采购定单"
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
               TabIndex        =   12
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
               Left            =   3780
               TabIndex        =   11
               Top             =   660
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
               TabIndex        =   10
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
               Left            =   3780
               TabIndex        =   9
               Top             =   1200
               Width           =   1155
               _Version        =   1048578
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "运费结算方式:"
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   255
               Left            =   7500
               TabIndex        =   8
               Top             =   630
               Width           =   495
               _Version        =   1048578
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "备注:"
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
            Height          =   5595
            Left            =   45
            TabIndex        =   14
            Top             =   1890
            Width           =   12810
            _ExtentX        =   22595
            _ExtentY        =   9869
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
         Begin TrueOleDBGrid80.TDBGrid TDBGrid3 
            Height          =   1785
            Left            =   8355
            TabIndex        =   16
            Top             =   45
            Width           =   2280
            _ExtentX        =   4022
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
      End
   End
End
Attribute VB_Name = "frmOriginalcontract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private id As String
Private rss As RecordSet
Private Originalsuppliers As String
Public mvarObjectID As String
Private UserName As String
Private theBLTool As New clsAutoCreateBL
Private Const theObjectID As String = "12B004"
Private theRsFreight As RecordSet

'运费主表的id
Private Costid As String

Private szCodeID As String

Private Function GetCodeID() As String
    GetCodeID = theBLTool.GetFrameCodeDetail(theObjectID)
End Function
Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property

Private Sub InitRsFreight()
    Set theRsFreight = New RecordSet
    theRsFreight.Fields.Append "B_ID", adInteger
    theRsFreight.Fields.Append "B_CodeID", adVarChar, 100
    'theRsFreight.Fields.Append "B_Qty", adDouble
    theRsFreight.Open
    
    
    TDBGrid3.DataSource = theRsFreight
    setgrid3
End Sub
Private Sub setgrid3()
    TDBGrid3.Columns("B_ID").width = 0
    TDBGrid3.Columns("B_ID").Visible = False
    TDBGrid3.Columns("B_ID").AllowSizing = False
    TDBGrid3.Columns("B_CodeID").Caption = "单据编号"
    TDBGrid3.MarqueeStyle = dbgHighlightRow
    TDBGrid3.HoldFields
End Sub
'添加一个采购入库单的运费
Private Sub AddFreight(ByVal vID As Long, ByVal vCodeID As String)
    theRsFreight.addnew
    theRsFreight!B_CodeID = vCodeID
    theRsFreight!B_ID = vID
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    
    Select Case Tool.name
        Case "保存"
            save
        Case "新增"
            AddBill
        Case "退出"
            eunload
        Case "新增"
            add
        Case "删除"
            del
        Case "第一单"
            movefrist
        Case "前一单"
            MovePreview
        Case "后一单"
            movenext
        Case "最后单"
            movelast
        Case "保存并打印"
            Saveprint
    End Select
End Sub

Private Sub Saveprint()
     If Trim(FlatEdit3.Text) = "" Then
        MsgBox "供应商不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(ComboBox2.Text) = "" Then
        MsgBox "交货方式不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(ComboBox3.Text) = "" Then
        MsgBox "交易方式不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If TDBGrid2.ApproxCount <= 0 Then
         MsgBox "明细表没有数据，不能保存", vbInformation, "提示"
         Exit Sub
    End If
    Dim sql1 As String
    Dim rs As New RecordSet
    sql1 = "select * from G_BillYarn where B_id='" & id & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If validation = True Then
            Update
            If id > 0 Then
                AddFreight id, szCodeID
            End If
        End If
    Else
        Dim sql As String
        Dim f As String
        f = Format(DTPicker1.Value, "YYYY-MM-dd")
        sql = "insert into G_BillYarn (B_ID,B_CodeID,B_BillType,B_Date,B_ContactCom,B_delivery,B_Balance,B_Memo,B_UserName) values('" & id & "','" & GetCodeID & "','YARN08','" & f & "','" & Originalsuppliers & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & Trim(FlatEdit4.Text) & "','" & Gm.SysID.SystemUser & "')"
        Gm.cnnTool.cnn.Execute sql
        
        savedetail
        MsgBox "保存成功", vbInformation, "提示"
        szCodeID = GetCodeID
        AddFreight id, szCodeID
    End If




    Dim sql4 As String
    Dim rs4 As New RecordSet
    
    Dim frm1 As New frmModBLRPreviewOri
    

    sql4 = "exec usp_PrintYarnOrder '" & id & "'"
    Debug.Print sql
    rs4.Open sql4, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

    
    Set frm1.RecordSet = rs4.Clone
    frm1.ObjectID = "22B024"
    frm1.Show
    
    
    
End Sub

Private Sub AddBill()
       addnew
        FlatEdit3.Text = ""
        delivery
        ClearWay
        Originalsuppliers = ""
End Sub
Private Sub del()
    Dim sql As String
    Dim sql1 As String
    Dim sql2 As String
    Dim rs As New RecordSet
    If TDBGrid2.ApproxCount > 0 Then
        If IIf(IsNull(rss!B_itemid), "", rss!B_itemid) <> "" Then
            sql1 = "select * from G_BillDetailYarn where B_itemid='" & rss!B_itemid & "'"
            rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            If rs.RecordCount > 0 Then
                    If validation = True Then
                        sql2 = "delete from G_BillDetailYarn where B_itemid='" & rss!B_itemid & "'"
                        Gm.cnnTool.cnn.Execute sql2
                        rss.requery
                    End If
            Else
                sql = "delete from G_DraftBillDetailYarn where B_itemid='" & rss!B_itemid & "'"
                Gm.cnnTool.cnn.Execute sql
                 rss.requery
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    InitFrm
    addnew
'            TDBGrid1.HighlightRowStyle.BackColor = RGB(240, 240, 240)
'    TDBGrid1.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
'    TDBGrid1.HighlightRowStyle.ForeColor = &H80000008
'    TDBGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
'    TDBGrid1.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
End Sub
'运费
Private Sub freight()
    ComboBox1.AddItem "现金垫付"
    ComboBox1.AddItem "月结垫付"
    ComboBox1.AddItem "现付"
    ComboBox1.AddItem "月结"
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    delivery
    ClearWay
    DTPicker1.Value = Now
    InitRsFreight
    freight
End Sub
'交货方式
Private Sub delivery()
    ComboBox2.Clear
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
    ComboBox3.Clear
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

Private Sub PushButton1_Click()
    Dim frm1 As New frmPopupDanWei
    frm1.ContactType = "原料供应商"
    frm1.Show vbModal
    Originalsuppliers = frm1.Clientid
    FlatEdit3.Text = frm1.ClientName
    Unload frm1
End Sub
Private Sub add()
    Dim sql1 As String
    Dim rs As New RecordSet
    sql1 = "select * from G_BillYarn where B_id='" & id & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If validation = True Then
        
            Dim frm1 As New frmOriginalcontract_Edit
            frm1.id = id
            frm1.Show vbModal
            Unload frm1
            If frm1.bsave = True Then
               rss.requery
            End If
           End If
    End If
End Sub

Private Sub eunload()
    Dim sql As String
    Dim sql1 As String
    Dim rs As New RecordSet
    If TDBGrid2.ApproxCount > 0 Then
        sql1 = "select * from G_BilldetailYarn where B_ID='" & id & "'"
        rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        If rs.RecordCount <= 0 Then
            If MsgBox("是保存数据，否将删除数据", vbYesNo + vbDefaultButton2 + vbInformation, "提示") = vbNo Then
                sql = "delete from G_DraftBilldetailYarn where B_ID='" & id & "'"
                Gm.cnnTool.cnn.Execute sql
                Unload Me
            Else
                save
            End If
        End If
    
    End If
    Unload Me
End Sub
Private Sub TDBGrid2_DblClick()
    Dim frm1 As New frmOriginalcontract_Edit
    If TDBGrid2.ApproxCount > 0 Then
        frm1.itemid = IIf(IsNull(rss!B_itemid), "", rss!B_itemid)
        frm1.OriginalProduct = IIf(IsNull(rss!B_sid), "", rss!B_sid)
        frm1.FlatEdit1.Text = IIf(IsNull(rss!B_name), "", rss!B_name)
        frm1.FlatEdit2.Text = IIf(IsNull(rss!B_specification), "", rss!B_specification)
        frm1.FlatEdit3.Text = IIf(IsNull(rss!B_qty), "", rss!B_qty)
        frm1.FlatEdit4.Text = IIf(IsNull(rss!B_price), "", rss!B_price)
        frm1.FlatEdit5.Text = IIf(IsNull(rss!B_Sum), "", rss!B_Sum)
        frm1.DTPicker1.Value = IIf(IsNull(rss!B_DeliveryTime), "", rss!B_DeliveryTime)
        frm1.FlatEdit7.Text = IIf(IsNull(rss!B_MemoDetail), "", rss!B_MemoDetail)
    End If
    frm1.id = id
    frm1.Show vbModal
    Unload frm1
    If frm1.bsave = True Then
        rss.requery
    End If
End Sub

Private Sub addnew()
    Dim sql As String
    Dim sql1 As String
    Dim rs As New RecordSet
    sql = "select *from G_DraftBillYarn where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_datecreate = Now
    rs.Update
    id = rs!B_ID
    sql1 = "delete from G_DraftBillYarn where B_ID='" & id & "'"
  adddetail
End Sub

Private Sub adddetail()
    Dim sql As String
    Set rss = New RecordSet
'    sql = "select * from G_DraftBilldetailYarn where B_ID='" & id & "'"
    sql = "exec usp_selectOriginaldetail '" & id & "'"
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid2.DataSource = rss
    setgrid2
End Sub
Private Sub setgrid2()
    TDBGrid2.Columns("B_Name").Caption = "原料名称"
    TDBGrid2.Columns("B_specification").Caption = "规格"
    TDBGrid2.Columns("B_Qty").Caption = "数量"
    TDBGrid2.Columns("B_price").Caption = "单价"
    TDBGrid2.Columns("B_sum").Caption = "金额"
    TDBGrid2.Columns("B_DeliveryTime").Caption = "交期"
    TDBGrid2.Columns("B_MemoDetail").Caption = "备注"
    
    TDBGrid2.Columns("B_Name").Locked = True
    TDBGrid2.Columns("B_specification").Locked = True
    TDBGrid2.Columns("B_Qty").Locked = True
    TDBGrid2.Columns("B_sum").Locked = True
    TDBGrid2.Columns("B_DeliveryTime").Locked = True
    
    TDBGrid2.Columns("B_pactCode").width = 0
    TDBGrid2.Columns("B_pactCode").Visible = False
    TDBGrid2.Columns("B_pactCode").AllowSizing = False
    TDBGrid2.Columns("B_OrderCode").width = 0
    TDBGrid2.Columns("B_OrderCode").Visible = False
    TDBGrid2.Columns("B_OrderCode").AllowSizing = False
    TDBGrid2.Columns("B_itemid").width = 0
    TDBGrid2.Columns("B_itemid").Visible = False
    TDBGrid2.Columns("B_itemid").AllowSizing = False
    TDBGrid2.Columns("B_ID").width = 0
    TDBGrid2.Columns("B_ID").Visible = False
    TDBGrid2.Columns("B_ID").AllowSizing = False
    TDBGrid2.Columns("B_itemidb").width = 0
    TDBGrid2.Columns("B_itemidb").Visible = False
    TDBGrid2.Columns("B_itemidb").AllowSizing = False
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
Private Sub TDBGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar21.Bands("网格右键").PopupMenu
    End If
End Sub

Private Sub save()

    If Trim(FlatEdit3.Text) = "" Then
        MsgBox "供应商不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(ComboBox2.Text) = "" Then
        MsgBox "交货方式不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(ComboBox3.Text) = "" Then
        MsgBox "交易方式不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If TDBGrid2.ApproxCount <= 0 Then
         MsgBox "明细表没有数据，不能保存", vbInformation, "提示"
         Exit Sub
    End If
    Dim sql1 As String
    Dim rs As New RecordSet
    sql1 = "select * from G_BillYarn where B_id='" & id & "'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If validation = True Then
            Update
            If id > 0 Then
                AddFreight id, szCodeID
            End If
        End If
    Else
        Dim sql As String
        Dim f As String
        f = Format(DTPicker1.Value, "YYYY-MM-dd")
        sql = "insert into G_BillYarn (B_ID,B_CodeID,B_BillType,B_Date,B_ContactCom,B_delivery,B_Balance,B_Memo,B_UserName) values('" & id & "','" & GetCodeID & "','YARN08','" & f & "','" & Originalsuppliers & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & Trim(FlatEdit4.Text) & "','" & Gm.SysID.SystemUser & "')"
        Gm.cnnTool.cnn.Execute sql
        
        savedetail
        MsgBox "保存成功", vbInformation, "提示"
        szCodeID = GetCodeID
        AddFreight id, szCodeID
    End If
    
End Sub
Private Sub savedetail()
    Dim sql As String
     Do While Not rss.EOF
        sql = "exec usp_InsertOriginalOrder '" & rss!B_itemid & "','" & rss!B_ID & "','" & rss!B_sid & "','" & rss!B_specification & "'"
        sql = sql & ",'" & rss!B_qty & "','" & rss!B_price & "','" & rss!B_Sum & "','" & rss!B_DeliveryTime & "','" & rss!B_MemoDetail & "','" & rss!B_PactCode & "','" & rss!B_OrderCode & "'"
        Gm.cnnTool.cnn.Execute sql
        rss.movenext
    Loop
End Sub

Private Sub movefrist()
    
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 * from G_BillYarn a left outer join G_BillDetailYarn b on a.B_id=b.B_ID where a.B_BillType='YARN08' and b.B_OrderCode='888888'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount <= 0 Then
            MsgBox "当前没有数据", vbInformation, "提示"
     Else
            id = rs!B_ID
            openbill
     End If
    
End Sub
Private Sub MovePreview()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BillYarn a left outer join G_BillDetailYarn b on a.B_id=b.B_ID where a.B_BillType='YARN08' and b.B_OrderCode='888888' and a.B_ID<'" & id & "' Order by a.B_ID desc"
     Debug.Print sql
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      If rs.RecordCount <= 0 Then
        movefrist
        MsgBox "已经是第一单了", vbOKOnly + vbInformation, "提示"
        
     Else
        id = rs!B_ID
        openbill
     End If
End Sub
Private Sub movenext()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BillYarn a left outer join G_BillDetailYarn b on a.B_id=b.B_ID where a.B_BillType='YARN08' and b.B_OrderCode='888888' and a.B_ID>'" & id & "' Order by a.B_ID asc"
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      If rs.RecordCount <= 0 Then
         movelast
        MsgBox "已经是最后一单了", vbOKOnly + vbInformation, "提示"
       
    Else
        id = rs!B_ID
        openbill
    End If
End Sub
Private Sub movelast()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 * from G_BillYarn a left outer join G_BillDetailYarn b on a.B_id=b.B_ID where a.B_BillType='YARN08' and b.B_OrderCode='888888' Order by a.B_ID desc "
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount <= 0 Then
        MsgBox "当前没有任何数据！", vbOKOnly + vbInformation, "提示"
    Else
        id = rs!B_ID
        
        openbill   '根据全局变量Sid打开单据，主表明细表显示到UI的对应位置
    End If
End Sub
Private Sub openbill()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_BillYarn where B_ID='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        Originalsuppliers = rs!B_ContactCom
        FlatEdit3.Text = getClientName(IIf(IsNull(rs!B_ContactCom), "", rs!B_ContactCom))
        DTPicker1.Value = rs!B_Date
        FlatEdit4.Text = rs!B_memo
        ComboBox2.Text = rs!B_delivery
        ComboBox3.Text = rs!B_Balance
        UserName = rs!B_username
        szCodeID = rs!B_CodeID
        openBilldetail
    End If
End Sub
Private Sub openBilldetail()
    Dim sql As String
    Set rss = New RecordSet
'    sql = "select * from G_DraftBilldetailYarn where B_ID='" & id & "'"
    sql = "exec usp_selectOriginaldetail_Edit '" & id & "'"
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid2.DataSource = rss
    setgrid2
End Sub

Private Sub Update()
    Dim sql As String
    sql = "update G_BillYarn set B_contactCom='" & Originalsuppliers & "',B_date='" & DTPicker1.Value & "'"
    sql = sql & ",B_Memo='" & FlatEdit4.Text & "',B_delivery='" & ComboBox2.Text & "',B_Balance='" & ComboBox3.Text & "'  where B_id='" & id & "'"
    Gm.cnnTool.cnn.Execute sql
    updatedetail
End Sub
Private Sub updatedetail()
    Dim sql As String
        rss.MoveFirst
    Do While Not rss.EOF
        sql = "update G_BilldetailYarn set B_goodsid='" & rss!B_sid & "',B_specification='" & rss!B_specification & "'"
        sql = sql & ",B_Qty='" & rss!B_qty & "',B_price='" & rss!B_price & "',B_sum='" & rss!B_Sum & "'"
        sql = sql & ",B_DeliveryTime='" & rss!B_DeliveryTime & "',B_MemoDetail='" & rss!B_MemoDetail & "'"
        sql = sql & "where B_itemid='" & rss!B_itemid & "'"
        Gm.cnnTool.cnn.Execute sql
        rss.movenext
    Loop
     rss.MoveFirst
End Sub


Private Function getClientName(ByVal Clientid As String) As String
     Dim rs As New RecordSet
     Dim sql As String
     sql = "Select B_ClientName From G_ContactCompany Where B_ClientID='" & Clientid & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount > 0 Then
        getClientName = rs!B_ClientName
     Else
        getClientName = ""
     End If
End Function

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

'--------------------------------------------------------生成运费----------------------------

Private Sub PushButton2_Click()
    If Trim(FlatEdit1.Text) = "" Then
        MsgBox "车号不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If Trim(ComboBox1.Text) = "" Then
        MsgBox "运费不能为空", vbInformation, "提示"
        Exit Sub
    End If
    If TDBGrid3.ApproxCount <= 0 Then
        MsgBox "表中没有数据", vbInformation, "提示"
        Exit Sub
    End If
    saveCost
    saveCostdetail
    FlatEdit1.Text = ""
    ComboBox1.Text = ""
    InitRsFreight
    MsgBox "生成运费采购", vbInformation, "提示"
End Sub

Private Sub saveCost()
    Dim sql As String
    Dim sql1 As String
    Dim rs As New RecordSet
    Dim f As String
    f = Format(Now, "YYYY-MM-DD")
    sql1 = "select * from G_BillCost where 1=1"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.addnew
    rs!B_Date = f
    rs!B_username = Gm.SysID.SystemUser
    rs!B_PlateNumber = FlatEdit1.Text
    Debug.Print ComboBox1.Text
    rs!B_freight = Trim(ComboBox1.Text)
    rs.Update
    Costid = rs!B_ID
End Sub

Private Sub saveCostdetail()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select *from G_BilldetailCost where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    theRsFreight.MoveFirst
    Do While Not theRsFreight.EOF
        rs.addnew
        rs!B_ID = Costid
        rs!B_YPID = theRsFreight!B_ID
        rs.Update
        theRsFreight.movenext
    Loop
End Sub

