VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmSchedule 
   Caption         =   "订单进度表"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11520
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
   ScaleHeight     =   7245
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11520
      _LayoutVersion  =   1
      _ExtentX        =   20320
      _ExtentY        =   12779
      _DataPath       =   ""
      Bands           =   "frmSchedule.frx":0000
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
         _GridInfo       =   $"frmSchedule.frx":1AE6
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   4965
            Left            =   30
            TabIndex        =   12
            Top             =   1500
            Width           =   9855
            _cx             =   17383
            _cy             =   8758
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   3263743
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "进度数字对比|进度颜色对比"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   6
            Position        =   0
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
            TabHeight       =   0
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Picture(0)      =   "frmSchedule.frx":1B6A
            Picture(1)      =   "frmSchedule.frx":2104
            Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
               Height          =   4530
               Left            =   15
               TabIndex        =   13
               Top             =   420
               Width           =   9825
               _ExtentX        =   17330
               _ExtentY        =   7990
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
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000002&"
               _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
               _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFFF0E1&"
               _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80000002&"
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
               Height          =   4530
               Left            =   10470
               TabIndex        =   14
               Top             =   420
               Width           =   9825
               _ExtentX        =   17330
               _ExtentY        =   7990
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
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000002&"
               _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
               _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFFF0E1&"
               _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80000002&"
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
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   4800
               TabIndex        =   11
               Top             =   960
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   4800
               TabIndex        =   3
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Format          =   231407617
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
               Format          =   231407617
               CurrentDate     =   43110
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2880
               TabIndex        =   5
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
               TabIndex        =   6
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
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   3840
               TabIndex        =   10
               Top             =   1020
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号:"
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   480
               TabIndex        =   9
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
               TabIndex        =   8
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
               TabIndex        =   7
               Top             =   1020
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "客        户:"
            End
         End
      End
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsdetail As RecordSet
Private rsdetail2 As RecordSet
Public mvarObjectID As String
Private Originalsuppliers As String '供应商的id
Private rss As RecordSet
Private rss2 As RecordSet
Private chose As String

Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property
'绑定草稿数据
Private Sub SetRs()
    Set rsdetail = New RecordSet
    rsdetail.Fields.Append "B_Pactid", adVarChar, 100
    rsdetail.Fields.Append "B_pactcode", adVarChar, 100
    rsdetail.Fields.Append "B_ClientName", adVarChar, 100
    
    rsdetail.Fields.Append "B_Date", adVarChar, 100
    rsdetail.Fields.Append "B_id", adVarChar, 100
    rsdetail.Fields.Append "B_YarnInsert", adVarChar, 100
    rsdetail.Fields.Append "B_YarnDelivery", adVarChar, 100
    rsdetail.Fields.Append "B_WhiteInsert", adVarChar, 100
    rsdetail.Fields.Append "B_WhiteDelivery", adVarChar, 100
    rsdetail.Fields.Append "B_DepartColor", adVarChar, 100
    rsdetail.Fields.Append "B_DepartFb", adVarChar, 100
    
    rsdetail.Fields.Append "B_DepartDJu", adVarChar, 100
    rsdetail.Fields.Append "B_ProcessInsert", adVarChar, 100
    rsdetail.Fields.Append "B_ProcessDelivery", adVarChar, 100
    rsdetail.Open
    TDBGrid1.DataSource = rsdetail
    setgrid
End Sub
'绑定草稿数据
Private Sub setRs2()
    Set rsdetail2 = New RecordSet
    rsdetail2.Fields.Append "B_Pactid", adVarChar, 100
    rsdetail2.Fields.Append "B_pactcode", adVarChar, 100
    rsdetail2.Fields.Append "B_ClientName", adVarChar, 100
    
    rsdetail2.Fields.Append "B_Date", adVarChar, 100
    rsdetail2.Fields.Append "B_id", adVarChar, 100
    rsdetail2.Fields.Append "B_YarnInsert", adVarChar, 100
    rsdetail2.Fields.Append "B_YarnDelivery", adVarChar, 100
    rsdetail2.Fields.Append "B_WhiteInsert", adVarChar, 100
    rsdetail2.Fields.Append "B_WhiteDelivery", adVarChar, 100
    rsdetail2.Fields.Append "B_DepartColor", adVarChar, 100
    rsdetail2.Fields.Append "B_DepartFb", adVarChar, 100
    
    rsdetail2.Fields.Append "B_DepartDJu", adVarChar, 100
    rsdetail2.Fields.Append "B_ProcessInsert", adVarChar, 100
    rsdetail2.Fields.Append "B_ProcessDelivery", adVarChar, 100
    rsdetail2.Open
    TDBGrid2.DataSource = rsdetail2
'    setgrid
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "查询"
            Grid
            grid2
        Case "退出"
            Unload Me
        Case "保存"
            save
        Case "设置数量"
            setnum
            
    End Select
End Sub



Private Sub Form_Load()
    InitFrm
'    setRs
    '打开窗体就执行查询
    Grid
    grid2

     TDBGrid1.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid1.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid1.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    
      TDBGrid2.HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid2.Splits(0).HighlightRowStyle.BackColor = RGB(240, 240, 240)
    TDBGrid2.HighlightRowStyle.ForeColor = &H80000008
    TDBGrid2.MarqueeStyle = dbgHighlightRowRaiseCell
    TDBGrid2.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell
    

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
    frm1.ContactType = "客户"
    frm1.Show vbModal
    Originalsuppliers = frm1.clientid
    FlatEdit3.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub Grid()
'    setRs
    Set rss = New RecordSet
    Dim sql As String
    Dim a As String
    Dim b As String
    
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    sql = "exec usp_schedule '" & a & "','" & b & "','" & Text1.Text & "','" & Originalsuppliers & "'"
    Debug.Print sql
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid1.DataSource = rss
'    openbill
    setgrid
End Sub
Private Sub grid2()
'    setRs2
    Set rss2 = New RecordSet
    Dim sql As String
    Dim a As String
    Dim b As String
    
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    sql = "exec usp_schedule_2 '" & a & "','" & b & "','" & Text1.Text & "','" & Originalsuppliers & "'"
    Debug.Print sql
    rss2.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid2.DataSource = rss2
    setgrid2
End Sub
Private Sub setgrid()
'    setGridShow
    
'    TDBGrid1.Columns("B_YarnInsert").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_YarnDelivery").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_WhiteInsert").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_WhiteDelivery").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_DepartFb").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_DepartColor").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_DepartDJu").NumberFormat = "0.0"
'    TDBGrid1.Columns("B_ProcessInsert").NumberFormat = "0.0"
    TDBGrid1.Columns("B_DepartFb").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_DepartDJu").ValueItems.Presentation = dbgCheckBox
    
    TDBGrid1.Columns("B_pactcode").Locked = True
    TDBGrid1.Columns("B_ClientName").Locked = True
    TDBGrid1.Columns("B_Date").Locked = True
    TDBGrid1.Columns("B_YarnInsert").Locked = True
    TDBGrid1.Columns("B_ProcessInsert").Locked = True
    TDBGrid1.Columns("B_ProcessDelivery").Locked = True
    
    TDBGrid1.Columns("B_pactcode").Caption = "订单号"
    TDBGrid1.Columns("B_ClientName").Caption = "客户"
    TDBGrid1.Columns("B_Date").Caption = "订单日期"
    TDBGrid1.Columns("B_YarnInsert").Caption = "原料入库"
    TDBGrid1.Columns("B_YarnDelivery").Caption = "原料发货"
    TDBGrid1.Columns("B_WhiteInsert").Caption = "白坯入库"
    TDBGrid1.Columns("B_WhiteDelivery").Caption = "白坯发货"
    TDBGrid1.Columns("B_DepartFb").Caption = "打样/制版"
    TDBGrid1.Columns("B_DepartFb").Locked = False
    TDBGrid1.Columns("B_DepartColor").Locked = False
    TDBGrid1.Columns("B_DepartDJu").Locked = False
    
    TDBGrid1.Columns("B_DepartColor").Caption = "计划"
    TDBGrid1.Columns("B_DepartDJu").Caption = "染色/印花"
    TDBGrid1.Columns("B_ProcessInsert").Caption = "深加工入库"
    TDBGrid1.Columns("B_ProcessDelivery").Caption = "深加工发货"

    TDBGrid1.Columns("B_id").Visible = False
    TDBGrid1.Columns("B_id").Locked = True
    TDBGrid1.Columns("B_id").AllowSizing = False
        TDBGrid1.Columns("B_Pactid").Visible = False
    TDBGrid1.Columns("B_Pactid").Locked = True
    TDBGrid1.Columns("B_Pactid").AllowSizing = False
    
    TDBGrid1.Columns("B_YarnDelivery").Visible = False
    TDBGrid1.Columns("B_YarnDelivery").Locked = True
    TDBGrid1.Columns("B_YarnDelivery").AllowSizing = False
    TDBGrid1.Columns("B_WhiteDelivery").Visible = False
    TDBGrid1.Columns("B_WhiteDelivery").Locked = True
    TDBGrid1.Columns("B_WhiteDelivery").AllowSizing = False
    
       TDBGrid1.Columns("B_pactcode").width = 900
    TDBGrid1.Columns("B_ClientName").width = 3000
    TDBGrid1.Columns("B_Date").width = 1200
    TDBGrid1.Columns("B_YarnInsert").width = 1200
    TDBGrid1.Columns("B_YarnDelivery").width = 1200
    TDBGrid1.Columns("B_WhiteInsert").width = 1200
    TDBGrid1.Columns("B_WhiteDelivery").width = 1200
    TDBGrid1.Columns("B_DepartFb").width = 1200
    TDBGrid1.Columns("B_DepartColor").width = 1200
    TDBGrid1.Columns("B_DepartDJu").width = 1200
    TDBGrid1.Columns("B_ProcessInsert").width = 1400
    TDBGrid1.Columns("B_ProcessDelivery").width = 1400
    TDBGrid1.HoldFields
'    TDBGrid1.MarqueeStyle = dbgHighlightRow
'    sumall
End Sub
Private Sub setgrid2()
'    setGridShow
    
    TDBGrid2.Columns("B_YarnInsert").NumberFormat = "0.0"
    TDBGrid2.Columns("B_YarnDelivery").NumberFormat = "0.0"
    TDBGrid2.Columns("B_WhiteInsert").NumberFormat = "0.0"
    TDBGrid2.Columns("B_WhiteDelivery").NumberFormat = "0.0"
    TDBGrid2.Columns("B_DepartFb").NumberFormat = "0.0"
    TDBGrid2.Columns("B_DepartColor").NumberFormat = "0.0"
    TDBGrid2.Columns("B_DepartDJu").NumberFormat = "0.0"
    TDBGrid2.Columns("B_ProcessInsert").NumberFormat = "0.0"
    TDBGrid2.Columns("B_pactcode").Locked = True
    TDBGrid2.Columns("B_ClientName").Locked = True
    TDBGrid2.Columns("B_Date").Locked = True
    
    TDBGrid2.Columns("B_pactcode").Caption = "订单号"
    TDBGrid2.Columns("B_ClientName").Caption = "客户"
    TDBGrid2.Columns("B_Date").Caption = "订单日期"
    TDBGrid2.Columns("B_YarnInsert").Caption = "原料入库"
    TDBGrid2.Columns("B_YarnDelivery").Caption = "原料发货"
    TDBGrid2.Columns("B_WhiteInsert").Caption = "白坯入库"
    TDBGrid2.Columns("B_WhiteDelivery").Caption = "白坯发货"
    TDBGrid2.Columns("B_DepartFb").Caption = "打样/制版"
    TDBGrid2.Columns("B_DepartColor").Caption = "计划"
    TDBGrid2.Columns("B_DepartDJu").Caption = "染色/印花"
    TDBGrid2.Columns("B_ProcessInsert").Caption = "深加工入库"
    TDBGrid2.Columns("B_ProcessDelivery").Caption = "深加工发货"
    
    TDBGrid2.Columns("B_pactcode").width = 900
    TDBGrid2.Columns("B_ClientName").width = 3000
    TDBGrid2.Columns("B_Date").width = 1200
    TDBGrid2.Columns("B_YarnInsert").width = 900
    TDBGrid2.Columns("B_YarnDelivery").width = 900
    TDBGrid2.Columns("B_WhiteInsert").width = 900
    TDBGrid2.Columns("B_WhiteDelivery").width = 900
    TDBGrid2.Columns("B_DepartFb").width = 900
    TDBGrid2.Columns("B_DepartColor").width = 900
    TDBGrid2.Columns("B_DepartDJu").width = 900
    TDBGrid2.Columns("B_ProcessInsert").width = 1400
    TDBGrid2.Columns("B_ProcessDelivery").width = 1400

    TDBGrid2.Columns("B_id").Visible = False
    TDBGrid2.Columns("B_id").Locked = True
    TDBGrid2.Columns("B_id").AllowSizing = False
        TDBGrid2.Columns("B_Pactid").Visible = False
    TDBGrid2.Columns("B_Pactid").Locked = True
    TDBGrid2.Columns("B_Pactid").AllowSizing = False
        TDBGrid2.Columns("B_YarnDelivery").Visible = False
    TDBGrid2.Columns("B_YarnDelivery").Locked = True
    TDBGrid2.Columns("B_YarnDelivery").AllowSizing = False
    TDBGrid2.Columns("B_WhiteDelivery").Visible = False
    TDBGrid2.Columns("B_WhiteDelivery").Locked = True
    TDBGrid2.Columns("B_WhiteDelivery").AllowSizing = False
      bianse
    TDBGrid2.HoldFields
    TDBGrid2.MarqueeStyle = dbgHighlightRow
    
    
'    sumall
End Sub
Private Sub setGridShow()
    Dim cls1 As New clsGridShow
    With cls1
        .ObjectID = "11S040"
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
                strSQL = "update G_BLSField set B_GridWidth='" & dWidth & "' where B_ObjectID='11S040' and B_FieldName='" & szFieldName & "'"
                Gm.cnnTool.cnn.Execute strSQL
            End If
        End If
    Next
End Sub

Private Sub sumall()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim rs As New RecordSet
    Dim a As Double
    Dim b As Double
    Dim f As Double
    Dim c As String
    Dim d As String
    
    a = 0
    b = 0
    f = 0
    TDBGrid1.Columns("B_XHBL").FooterText = "合计"
    If rss.RecordCount <= 0 Then
        TDBGrid1.Columns("B_ps").FooterText = "" & a & ""
        TDBGrid1.Columns("B_KG").FooterText = "" & b & ""
        TDBGrid1.Columns("B_meter").FooterText = "" & f & ""
    End If
    
    
    Set rs = rss.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_ps), 0, rs!B_ps)
        b = b + IIf(IsNull(rs!B_kg), 0, rs!B_kg)
        f = f + IIf(IsNull(rs!B_meter), 0, rs!B_meter)
        rs.movenext
    Loop
    c = Format(a, "0.0")
'    d = Format(b, "0.00")
    TDBGrid1.Columns("B_ps").FooterText = "" & c & ""
    TDBGrid1.Columns("B_kg").FooterText = "" & b & ""
    TDBGrid1.Columns("B_meter").FooterText = "" & f & ""
End Sub

Private Sub save()
    Dim sql As String
    Dim rs As New RecordSet
    rsdetail.MoveFirst
    
    Do While Not rsdetail.EOF
        If Len(IIf(IsNull(rsdetail!B_id), "", rsdetail!B_id)) <= 0 Then
            sql = "exec usp_scheduleinsert '" & rsdetail!B_Pactid & "','" & rsdetail!B_YarnInsert & "','" & rsdetail!B_YarnDelivery & "','" & rsdetail!B_WhiteInsert & "','" & rsdetail!B_WhiteDelivery & "',"
            sql = sql & "'" & rsdetail!B_DepartFb & "','" & rsdetail!B_DepartColor & "','" & rsdetail!B_DepartDJu & "','" & rsdetail!B_ProcessInsert & "'"
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        Else
            sql = "exec usp_schedule_update '" & rsdetail!B_id & "','" & rsdetail!B_YarnInsert & "','" & rsdetail!B_YarnDelivery & "','" & rsdetail!B_WhiteInsert & "','" & rsdetail!B_WhiteDelivery & "',"
            sql = sql & "'" & rsdetail!B_DepartFb & "','" & rsdetail!B_DepartColor & "','" & rsdetail!B_DepartDJu & "','" & rsdetail!B_ProcessInsert & "'"
            Debug.Print sql
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        End If
        rsdetail.movenext
    Loop
    SetRs
    Grid
    
End Sub

Private Sub openbill()
    Do While Not rss.EOF
          rsdetail.AddNew
            rsdetail!B_Pactid = rss!B_Pactid
           rsdetail!B_PactCode = rss!B_PactCode
            rsdetail!B_ClientName = rss!B_ClientName
            
            rsdetail!B_Date = rss!B_Date
            rsdetail!B_id = IIf(IsNull(rss!B_id), "", rss!B_id)
            rsdetail!B_YarnInsert = IIf(IsNull(rss!B_YarnInsert), "", rss!B_YarnInsert)
            rsdetail!B_YarnDelivery = IIf(IsNull(rss!B_YarnDelivery), "", rss!B_YarnDelivery)
            rsdetail!B_WhiteInsert = IIf(IsNull(rss!B_WhiteInsert), "", rss!B_WhiteInsert)
            rsdetail!B_WhiteDelivery = IIf(IsNull(rss!B_WhiteDelivery), "", rss!B_WhiteDelivery)
            rsdetail!B_DepartFb = IIf(IsNull(rss!B_DepartFb), "", rss!B_DepartFb)
            rsdetail!B_DepartColor = IIf(IsNull(rss!B_DepartColor), "", rss!B_DepartColor)
            rsdetail!B_DepartDJu = IIf(IsNull(rss!B_DepartDJu), "", rss!B_DepartDJu)
            rsdetail!B_ProcessInsert = IIf(IsNull(rss!B_ProcessInsert), "", rss!B_ProcessInsert)
          rsdetail.Update
          rss.movenext
    Loop
    If TDBGrid1.ApproxCount > 0 Then
        rsdetail.MoveFirst
    End If
End Sub

Private Sub setnum()
    Dim a As Long
    Dim b As String
    Dim bookmark As Long
    bookmark = TDBGrid1.bookmark
    
    Dim sql As String
    Dim rs As New RecordSet
    If C1Tab1.CurrTab <> 0 Then
        Exit Sub
    End If
    
    If TDBGrid1.Col <= 7 Or TDBGrid1.Col > 9 Then
         Exit Sub
    End If
 
    
    
    Dim frm1 As New frmSchedule_Edit
    frm1.Show vbModal
    If frm1.bool = False Then
        Exit Sub
    End If
    a = frm1.FlatEdit2.Text
    Unload frm1
    b = TDBGrid1.Columns(TDBGrid1.Col).DataField
    Debug.Print b

    If Len(IIf(IsNull(rss!B_id), "", rss!B_id)) <= 0 Then
        sql = "insert into G_schedule (B_orderid," & b & ") values('" & rss!B_Pactid & "','" & a & "')"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
    
    
'        sql = "exec usp_scheduleinsert '" & rsdetail!B_Pactid & "','" & rsdetail!B_YarnInsert & "','" & rsdetail!B_YarnDelivery & "','" & rsdetail!B_WhiteInsert & "','" & rsdetail!B_WhiteDelivery & "',"
'        sql = sql & "'" & rsdetail!B_DepartFb & "','" & rsdetail!B_DepartColor & "','" & rsdetail!B_DepartDJu & "','" & rsdetail!B_ProcessInsert & "'"
'        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Else
        sql = "update G_schedule set " & b & "='" & a & "' where B_id='" & rss!B_id & "'"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
'        sql = "exec usp_schedule_update '" & rsdetail!B_id & "','" & rsdetail!B_YarnInsert & "','" & rsdetail!B_YarnDelivery & "','" & rsdetail!B_WhiteInsert & "','" & rsdetail!B_WhiteDelivery & "',"
'        sql = sql & "'" & rsdetail!B_DepartFb & "','" & rsdetail!B_DepartColor & "','" & rsdetail!B_DepartDJu & "','" & rsdetail!B_ProcessInsert & "'"
'        Debug.Print sql
'        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    End If
'    setRs
    Grid
    grid2
    TDBGrid1.bookmark = bookmark
End Sub
'进行修改排版和染色
Private Sub setlogo()
    Dim bookmark As Long
     Dim a As Long
    Dim b As String
    Dim sql As String
    Dim rs As New RecordSet
    bookmark = TDBGrid1.bookmark
'    Dim frm1 As New frmSchedule_Edit2
'
'    frm1.Show vbModal
'    If frm1.bool = False Then
'        Exit Sub
'    End If
'    a = Abs(Val(frm1.Check1.Value))

    If TDBGrid1.Columns(TDBGrid1.Col).Value = 1 Then
        a = 0
    Else
        a = 1
    End If



'    Unload frm1
    b = TDBGrid1.Columns(TDBGrid1.Col).DataField
    If Len(IIf(IsNull(rss!B_id), "", rss!B_id)) <= 0 Then
        sql = "insert into G_schedule (B_orderid," & b & ") values('" & rss!B_Pactid & "','" & a & "')"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
    Else
        sql = "update G_schedule set " & b & "='" & a & "' where B_id='" & rss!B_id & "'"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
'Dim m As Long
'Dim n As Long
'm = 1
'n = 1
' If Abs(rss!B_DepartFb) = 1 Then
'       m = 0
' End If
'  If Abs(rss!B_DepartDJu) = 1 Then
'     n = 0
' End If
'
'        If Len(IIf(IsNull(rss!B_id), "", rss!B_id)) <= 0 Then
''        sql = "insert into G_schedule (B_orderid,B_DepartFb,B_DepartDJu) values('" & rss!B_Pactid & "','" & Abs(rss!B_DepartFb) & "','" & Abs(rss!B_DepartDJu) & "')"
'        sql = "insert into G_schedule (B_orderid,B_DepartFb,B_DepartDJu) values('" & rss!B_Pactid & "','" & m & "','" & n & "')"
'        Debug.Print sql
'        Gm.cnnTool.cnn.Execute sql
'    Else
'        sql = "update G_schedule SET B_DepartFb='" & m & "',B_DepartDJu='" & n & "'  where B_id='" & rss!B_id & "'"
'        Debug.Print sql
'        Gm.cnnTool.cnn.Execute sql
'    End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Grid
    grid2
    TDBGrid1.bookmark = bookmark
End Sub

Private Sub bianse()
    Dim i As Long
    For i = 0 To TDBGrid2.Columns.Count - 1
            TDBGrid2.Columns(i).FetchStyle = True
    Next
End Sub

Private Sub TDBGrid1_Click()
    If TDBGrid1.Col <= 11 And TDBGrid1.Col > 9 Then
                'setlogo
                Exit Sub
    End If
End Sub

Private Sub TDBGrid1_ColEdit(ByVal ColIndex As Integer)
    If TDBGrid1.Col <= 11 And TDBGrid1.Col > 9 Then
        setlogo
    End If
End Sub

Private Sub TDBGrid2_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, _
    bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    
    Dim i As Long
    Dim j As Long
    Dim m_Num As Long
    
    
    '需要做的工序 -已刷卡
'    If CountChinese(TDBGrid2.Columns(Col).DataField) > 0 Then
        If Val(TDBGrid2.Columns(Col).CellValue(bookmark)) > 0 And Val(TDBGrid2.Columns(Col).CellValue(bookmark)) < 1 Then
            'If TDBGrid2.Col > 4 Then
            If Col > 4 Then
                CellStyle.BackColor = vbGreen
                
                CellStyle.ForeColor = CellStyle.BackColor
            End If
            'End If
        End If
'    End If

    '需要做的工序 - 未刷卡
'    If CountChinese(TDBGrid2.Columns(Col).DataField) > 0 Then
    
    If Val(TDBGrid2.Columns(Col).CellValue(bookmark)) = 0 Then
        'If TDBGrid2.Col > 4 Then
        If Col > 4 Then
            CellStyle.BackColor = vbRed
            
            CellStyle.ForeColor = CellStyle.BackColor
        End If
        'End If
    End If
'    End If
    
    
    Debug.Print Val(TDBGrid2.Columns(Col).CellValue(bookmark))
    '不需要做的工序
'    If CountChinese(TDBGrid2.Columns(Col).DataField) > 0 Then
    If Val(TDBGrid2.Columns(Col).CellValue(bookmark)) >= 1 Then
        'If TDBGrid2.Col > 4 Then
        If Col > 4 Then
            CellStyle.BackColor = &HFFFF&
            
            CellStyle.ForeColor = CellStyle.BackColor
        End If
        'End If
    End If
    
    
'    End If
End Sub
'Private Function CountChinese(ByVal txtStr As String) As Long
'    Dim i As Long
'    Dim alls As Long
'    For i = 1 To Len(txtStr)
'        If Asc(Mid(txtStr, i, 1)) < 0 Then
'            alls = alls + 1
'        End If
'    Next i
'
'    CountChinese = alls
'End Function


