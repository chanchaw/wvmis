VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmDingDanSelect 
   Caption         =   "订单查询"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDingDanSelect.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11595
      _LayoutVersion  =   1
      _ExtentX        =   20452
      _ExtentY        =   13785
      _DataPath       =   ""
      Bands           =   "frmDingDanSelect.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6735
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   11175
         _cx             =   19711
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
         GridRows        =   4
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmDingDanSelect.frx":2940
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1185
            Left            =   30
            ScaleHeight     =   1185
            ScaleWidth      =   11115
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   11115
            Begin VB.CommandButton Command1 
               Caption         =   "修复条码"
               Height          =   480
               Left            =   4800
               TabIndex        =   17
               Top             =   247
               Width           =   1110
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   375
               Left            =   1080
               TabIndex        =   3
               Top             =   300
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
            Begin XtremeSuiteControls.Label Label9 
               Height          =   255
               Left            =   360
               TabIndex        =   4
               Top             =   360
               Width           =   615
               _Version        =   1048578
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号:"
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
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   5460
            Left            =   30
            TabIndex        =   5
            Top             =   1245
            Width           =   11115
            _cx             =   19606
            _cy             =   9631
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
            Appearance      =   0
            MousePointer    =   0
            Version         =   800
            BackColor       =   13557726
            ForeColor       =   -2147483630
            FrontTabColor   =   3263743
            BackTabColor    =   8355711
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "色布计划|白坯计划"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   6
            Position        =   5
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   -1  'True
            TabsPerPage     =   0
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   1000
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Picture(0)      =   "frmDingDanSelect.frx":29C3
            Picture(1)      =   "frmDingDanSelect.frx":2D5D
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   5430
               Left            =   1020
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   15
               Width           =   10080
               _cx             =   17780
               _cy             =   9578
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
               ChildSpacing    =   0
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
               GridRows        =   6
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmDingDanSelect.frx":32F7
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic C1Elastic9 
                  Height          =   5370
                  Left            =   30
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   10020
                  _cx             =   17674
                  _cy             =   9472
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
                  ChildSpacing    =   0
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
                  GridRows        =   6
                  GridCols        =   4
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmDingDanSelect.frx":3390
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin TrueOleDBGrid80.TDBGrid TDBGrid3 
                     Height          =   4935
                     Left            =   30
                     TabIndex        =   8
                     Top             =   405
                     Width           =   9960
                     _ExtentX        =   17568
                     _ExtentY        =   8705
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
                     _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                     _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
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
                     _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgpicMode=2,.bgbmp=2"
                     _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bgpicMode=2,.bgbmp=1"
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
                  Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar24 
                     Height          =   375
                     Left            =   30
                     TabIndex        =   9
                     Top             =   30
                     Width           =   9960
                     _LayoutVersion  =   1
                     _ExtentX        =   17568
                     _ExtentY        =   661
                     _DataPath       =   ""
                     Bands           =   "frmDingDanSelect.frx":3428
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   5430
               Left            =   12735
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   15
               Width           =   10080
               _cx             =   17780
               _cy             =   9578
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
               _GridInfo       =   $"frmDingDanSelect.frx":47B2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   5370
                  Left            =   30
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   10020
                  _cx             =   17674
                  _cy             =   9472
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
                  ChildSpacing    =   0
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
                  GridRows        =   6
                  GridCols        =   4
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmDingDanSelect.frx":4832
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar23 
                     Height          =   360
                     Left            =   30
                     TabIndex        =   12
                     Top             =   30
                     Width           =   9960
                     _LayoutVersion  =   1
                     _ExtentX        =   17568
                     _ExtentY        =   635
                     _DataPath       =   ""
                     Bands           =   "frmDingDanSelect.frx":48CA
                  End
                  Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
                     Height          =   4950
                     Left            =   30
                     TabIndex        =   13
                     Top             =   390
                     Width           =   9960
                     _ExtentX        =   17568
                     _ExtentY        =   8731
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
                     _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                     _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                     _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
                     _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
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
                     _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgpicMode=2,.bgbmp=1"
                     _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                     _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                     _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                     _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                     _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgpicMode=2"
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
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   5430
               Left            =   13035
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   15
               Width           =   10080
               _cx             =   17780
               _cy             =   9578
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
               ChildSpacing    =   0
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
               GridRows        =   6
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmDingDanSelect.frx":5BCC
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar25 
                  Height          =   420
                  Left            =   30
                  TabIndex        =   15
                  Top             =   30
                  Width           =   10020
                  _LayoutVersion  =   1
                  _ExtentX        =   17674
                  _ExtentY        =   741
                  _DataPath       =   ""
                  Bands           =   "frmDingDanSelect.frx":5C63
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid5 
                  Height          =   4950
                  Left            =   30
                  TabIndex        =   16
                  Top             =   450
                  Width           =   10020
                  _ExtentX        =   17674
                  _ExtentY        =   8731
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
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
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
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgpicMode=2,.bgbmp=1"
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
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmDingDanSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'网格1记录集
Private rss As RecordSet
'网格2记录集
Private rss1 As RecordSet

Public mvarObjectID As String

Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property


Private Sub Command1_Click()
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As RecordSet
    Dim lIncr As String
    Dim szBC13 As String
    sql = "SELECT G_BilldetailColor.* FROM G_BilldetailColor LEFT OUTER JOIN G_BillColor ON G_BilldetailColor.B_ID=G_BillColor.B_ID"
    sql = sql & " WHERE   G_BillColor.B_Billtype='COL01' AND isnull(B_BC13,'')=''"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
                '获取最新的一个条码的自增数字
            lIncr = GetNewBCIncr
            szBC13 = GetBC13(FillGetBC12(lIncr))
            Set rs1 = New RecordSet
            sql1 = "update G_BilldetailColor set B_BCIncr='" & lIncr & "',B_BC13='" & szBC13 & "' where B_itemid='" & rs!B_ItemID & "'"
             rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs.movenext
        Loop

    End If
    
End Sub

Private Sub FlatEdit3_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
           Grid
    End Select
End Sub

Private Sub Form_Load()
    InitFrm
    Grid
    gridwhite
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "查询"
            Grid
        Case "合同明细打印"
'            PrintDetail
        Case "退出"
            Unload Me
        Case "设置转发"
            setnote
        Case "设置打卷"
            setdj
        Case "设置色布打卷样式"
            ColorDaJuan
    End Select
End Sub

'------------------------------------------色布计划--------------------------------------
Private Sub Grid()
    Dim sql As String
    Set rss = New RecordSet
    Dim sql1 As String
    Dim rs As New RecordSet
    sql = "exec usp_SelectDingDanColor '" & Trim(FlatEdit3.Text) & "'"
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid3.DataSource = rss
    setgrid
    gridwhite
    sql1 = "select * from G_BillDetailOrder where B_OrderCode like '%'+ '" & Trim(FlatEdit3.Text) & "' +'%'"
    rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount <= 0 Then
        MsgBox "订单号不存在", vbInformation, "提示"
    Else
        If rss.RecordCount <= 0 And rss1.RecordCount <= 0 Then
             MsgBox "色布计划和白坯计划不存在", vbInformation, "提示"
        Else
            If rss.RecordCount > 0 And rss1.RecordCount <= 0 Then
                MsgBox "白坯计划不存在", vbInformation, "提示"
            End If
            If rss.RecordCount <= 0 And rss1.RecordCount > 0 Then
                MsgBox "色布计划不存在", vbInformation, "提示"
            End If
        End If
    End If
End Sub


'设置网格1的样式
Private Sub setgrid()
    
    TDBGrid3.Columns("B_ClientName").Caption = "客户"
    TDBGrid3.Columns("B_alias").Caption = "别名"
    TDBGrid3.Columns("B_PactCode").Caption = "合同号"
    TDBGrid3.Columns("B_ItemIDB").Caption = "订单号"
    TDBGrid3.Columns("B_GoodsNameAlias").Caption = "色布名称"
    TDBGrid3.Columns("B_orderWidth").Caption = "订单门幅"
    TDBGrid3.Columns("B_orderWeight").Caption = "订单克重"
    TDBGrid3.Columns("B_width").Caption = "染厂门幅"
    TDBGrid3.Columns("B_weight").Caption = "染厂克重"
    TDBGrid3.Columns("B_practiceCast").Caption = "数量"
    TDBGrid3.Columns("B_Hex").Caption = "颜色示例"
    TDBGrid3.Columns("B_producer").Caption = "花型"
    TDBGrid3.Columns("B_SeHao").Caption = "花号"
    TDBGrid3.Columns("B_orderColor").Caption = "颜色"
    
    TDBGrid3.Columns("B_Hex").width = 900
    TDBGrid3.Columns("B_ClientName").width = 3000
    TDBGrid3.Columns("B_alias").width = 1500
    TDBGrid3.Columns("B_width").width = 1400
    TDBGrid3.Columns("B_weight").width = 1500
    TDBGrid3.Columns("B_practiceCast").width = 1000
    TDBGrid3.Columns("B_producer").width = 1000
    TDBGrid3.Columns("B_SeHao").width = 1000
    TDBGrid3.Columns("B_orderColor").width = 800
    TDBGrid3.Columns("B_PactCode").width = 1400
    TDBGrid3.Columns("B_ItemIDB").width = 1400
    
    TDBGrid3.Columns("B_itemid").Visible = False
    TDBGrid3.Columns("B_itemid").AllowSizing = False
    TDBGrid3.Columns("B_itemid").Locked = True
    
    TDBGrid3.Columns("B_BelongOrderID").Visible = False
    TDBGrid3.Columns("B_BelongOrderID").AllowSizing = False
    TDBGrid3.Columns("B_BelongOrderID").Locked = True
    
    TDBGrid3.Columns("B_flowCardprint").Visible = False
    TDBGrid3.Columns("B_flowCardprint").AllowSizing = False
    TDBGrid3.Columns("B_flowCardprint").Locked = True
    
    TDBGrid3.Columns("B_Hex").FetchStyle = True
    
    TDBGrid3.HoldFields
    TDBGrid3.MarqueeStyle = dbgHighlightRow
End Sub



Private Sub TDBGrid3_DblClick()
    Dim a As Long
    a = TDBGrid3.bookmark
    
    If TDBGrid3.ApproxCount <= 0 Then
        Exit Sub
    End If

    Dim frm1 As New frmDingDanSelect_Edit
    frm1.item = rss!B_ItemID
    frm1.itemidb = rss!B_ItemIDB
    frm1.Show vbModal
    Unload frm1
    rss.requery
   TDBGrid3.bookmark = a
End Sub


Private Sub TDBGrid3_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

'    Dim ys As RGB
    On Error Resume Next
    Debug.Print TDBGrid3.Columns("B_Hex").CellValue(bookmark)
    CellStyle.BackColor = TDBGrid3.Columns("B_Hex").CellValue(bookmark)
     CellStyle.ForeColor = TDBGrid3.Columns("B_Hex").CellValue(bookmark)
End Sub

Private Sub ActiveBar24_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "打印流传卡"
            printOne
        Case "打印全部流传卡"
            printall
    End Select
End Sub
Private Sub printOne()
     If TDBGrid3.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim rs1 As New RecordSet
    Dim sql1 As String
    sql1 = "SELECT * FROM  G_BillDetailColor WHERE B_ItemID='" & rss!B_ItemID & "' AND LEN(B_DaJuanGS)>0"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs1.RecordCount <= 0 Then
         MsgBox "没有设置色布打卷的样式！", vbInformation, "提示"
        Exit Sub
    End If
    
    Dim rs As New RecordSet
    Dim sql As String
    sql = "exec usp_printCard '" & rss!B_BelongOrderID & "','" & rss!B_ItemID & "','" & Gm.SysID.SystemUserName & "'"
    Debug.Print sql
    
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     Dim frm1 As New ActiveReport2
    frm1.itmeid = rss!B_ItemID
    frm1.flowCardprint = IIf(IsNull(rss!B_flowCardprint), 0, rss!B_flowCardprint)
    Set frm1.rs = rs.Clone
    frm1.Show vbModal
    rss.requery
End Sub

Private Sub printall()
      If TDBGrid3.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim rs As RecordSet
    Dim frm1 As ActiveReport2
    rss.MoveFirst
    Do While Not rss.EOF
        Set frm1 = New ActiveReport2
        Set rs = New RecordSet
        frm1.itmeid = rss!B_ItemID
         frm1.flowCardprint = IIf(IsNull(rss!B_flowCardprint), 0, rss!B_flowCardprint)
        Dim sql As String
        sql = "exec usp_printCard '" & rss!B_BelongOrderID & "','" & rss!B_ItemID & "','" & Gm.SysID.SystemUserName & "'"
        Debug.Print sql
         rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        Set frm1.rs = rs.Clone
        frm1.Show vbModal
        rss.movenext
    Loop
     rss.requery
End Sub




'----------------------------------------白坯计划--------------------------------------------

Private Sub gridwhite()
    Dim sql As String
    Set rss1 = New RecordSet
    sql = "exec usp_SelectDingDanwhite '" & Trim(FlatEdit3.Text) & "'"
    rss1.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid2.DataSource = rss1
    setgridwhite
    
End Sub

Private Sub setgridwhite()
    TDBGrid2.Columns("B_ClientName").Caption = "客户"
    TDBGrid2.Columns("B_PactCode").Caption = "合同号"
    TDBGrid2.Columns("B_ItemIDB").Caption = "订单号"
    TDBGrid2.Columns("B_Name").Caption = "白坯名称"
    TDBGrid2.Columns("B_Width").Caption = "门幅"
    TDBGrid2.Columns("B_UnitWeight").Caption = "克重"
    TDBGrid2.Columns("B_MaoHight").Caption = "毛高"
    TDBGrid2.Columns("B_CastQty").Caption = "投份"
    TDBGrid2.Columns("B_Maospecification").Caption = "毛丝规格"
    TDBGrid2.Columns("B_BOXQty").Caption = "数量"
    TDBGrid2.Columns("B_attention").Caption = "注意事项"
    
    TDBGrid2.Columns("B_ItemID").AllowSizing = False
    TDBGrid2.Columns("B_ItemID").Locked = True
    TDBGrid2.Columns("B_ItemID").Visible = False
    TDBGrid2.Columns("B_BelongOrderID").AllowSizing = False
    TDBGrid2.Columns("B_BelongOrderID").Locked = True
    TDBGrid2.Columns("B_BelongOrderID").Visible = False
    
    TDBGrid2.Columns("B_print").AllowSizing = False
    TDBGrid2.Columns("B_print").Locked = True
    TDBGrid2.Columns("B_print").Visible = False
    
    TDBGrid2.MarqueeStyle = dbgHighlightRow
    TDBGrid2.HoldFields
End Sub


Private Sub TDBGrid2_DblClick()
      If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If

    Dim frm1 As New frmDingDanSelect_Edit1
    frm1.item = rss1!B_ItemID
    frm1.itemidb = rss1!B_ItemIDB
    frm1.Show vbModal
    Unload frm1
    rss1.requery
End Sub
Private Sub ActiveBar23_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "打印当前订单"
            printwhite
        Case "打印全部订单"
            printwhiteAll
    End Select
End Sub
Private Sub printwhite()
    If rss1.RecordCount <= 0 Then
        Exit Sub
    End If
     Dim rs As New RecordSet
    Dim sql As String
    sql = "exec usp_printWhite '" & rss1!B_BelongOrderID & "','" & rss1!B_ItemID & "'"
    
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Dim frm1 As New ActiveReport8
    frm1.itmeid = rss1!B_ItemID
    frm1.flowCardprint = IIf(IsNull(rss1!B_print), 0, rss1!B_print)
    Set frm1.rs = rs.Clone
    
    frm1.Show vbModal
    rss1.requery
End Sub
Private Sub printwhiteAll()
     If rss1.RecordCount <= 0 Then
        Exit Sub
    End If
     Dim rs As RecordSet
    Dim frm1 As New ActiveReport8
    Dim sql As String
    Do While Not rss1.EOF
    Set rs = New RecordSet
    sql = "exec usp_printWhite '" & rss1!B_BelongOrderID & "','" & rss1!B_ItemID & "'"
    Debug.Print sql
    
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set frm1 = New ActiveReport8
    frm1.itmeid = rss1!B_ItemID
    frm1.flowCardprint = IIf(IsNull(rss1!B_print), 0, rss1!B_print)
    Set frm1.rs = rs.Clone
    frm1.Show vbModal
    rss1.movenext
    Loop
    rss1.requery
End Sub

Private Sub setnote()
    On Error Resume Next
    
    Dim a As Long
    a = TDBGrid3.bookmark
    
    If TDBGrid3.ApproxCount <= 0 Then
        Exit Sub
    End If

    Dim frm1 As New frmDingDanSelect_Edit
    frm1.item = rss!B_ItemID
    frm1.itemidb = rss!B_ItemIDB
    frm1.Show vbModal
    Unload frm1
    rss.requery
    TDBGrid3.bookmark = a
End Sub


'从表G_BillDetailColor获取当前最新一个条码的自增数字
Private Function GetNewBCIncr() As Long
    Dim rs As New RecordSet
    strSQL = "select top 1 * from G_BillDetailColor order by B_BCIncr desc"
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

Private Sub setdj()
    Dim sql As String
    Dim rs As New RecordSet
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    If C1Tab1.CurrTab <> 1 Then
        Exit Sub
    End If
    Dim frm1 As New frmDingDanSelect_Edit2
    frm1.Show vbModal
    If frm1.bool = False Then
        Exit Sub
    End If
    
    sql = "insert into G_WhiteCurly2(B_itemidb,B_BCitemid,B_DTRK,B_PS,B_KG,B_type,B_ZCP) "
    sql = sql & " values('" & rss1!B_ItemIDB & "','" & rss1!B_ItemID & "','" & frm1.DTPicker1.Value & "','" & frm1.FlatEdit3.Text & "','" & frm1.FlatEdit1.Text & "','" & frm1.ComboBox1.Text & "','" & frm1.ComboBox2.Text & "')"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Unload frm1
End Sub

Private Sub ColorDaJuan()
Dim frm1 As New frmDingDanSelect_Edit3
Dim m_DaJuanGS As String
Dim sql As String
Dim rs As New RecordSet
frm1.m_ID = rss!B_ItemID
frm1.Show vbModal
    If frm1.bsaved = True Then
        m_DaJuanGS = frm1.m_DaJuanGS
        
        sql = "UPDATE G_BillDetailColor SET B_DaJuanGS='" & m_DaJuanGS & "'  WHERE B_itemid='" & rss!B_ItemID & "'"
        Gm.cnnTool.cnn.Execute sql
    End If

Unload frm1

End Sub
