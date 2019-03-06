VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmProgress 
   Caption         =   "订单进度表"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "订单进度表"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _LayoutVersion  =   1
      _ExtentX        =   21828
      _ExtentY        =   14314
      _DataPath       =   ""
      Bands           =   "frmProgress.frx":038A
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5895
         Left            =   360
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   11055
         _cx             =   19500
         _cy             =   10398
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
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
         _GridInfo       =   $"frmProgress.frx":162E
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1425
            Left            =   30
            ScaleHeight     =   1425
            ScaleWidth      =   10995
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   30
            Width           =   10995
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   4320
               TabIndex        =   6
               Top             =   720
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   4320
               TabIndex        =   7
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Format          =   220135425
               CurrentDate     =   43110
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   1080
               TabIndex        =   8
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Format          =   220135425
               CurrentDate     =   43110
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   375
               Left            =   2640
               TabIndex        =   9
               Top             =   720
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
               Left            =   1080
               TabIndex        =   10
               Top             =   720
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
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Left            =   0
               TabIndex        =   14
               Top             =   780
               Width           =   1095
               _Version        =   1048578
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "染         厂:"
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Left            =   3360
               TabIndex        =   13
               Top             =   60
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "终止日期:"
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   255
               Left            =   0
               TabIndex        =   12
               Top             =   60
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "起始日期:"
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   3360
               TabIndex        =   11
               Top             =   780
               Width           =   855
               _Version        =   1048578
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号:"
            End
         End
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   4380
            Left            =   30
            TabIndex        =   2
            Top             =   1485
            Width           =   10995
            _cx             =   19394
            _cy             =   7726
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   134
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
            FrontTabForeColor=   0
            Caption         =   "数量进度|颜色进度"
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
            TabHeight       =   500
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   3855
               Left            =   15
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   510
               Width           =   10965
               _cx             =   19341
               _cy             =   6800
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   134
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
               _GridInfo       =   $"frmProgress.frx":16AE
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
                  Bindings        =   "frmProgress.frx":172E
                  Height          =   3795
                  Left            =   30
                  TabIndex        =   4
                  Top             =   30
                  Width           =   10905
                  _ExtentX        =   19235
                  _ExtentY        =   6694
                  _LayoutType     =   0
                  _RowHeight      =   31
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
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).DividerColor=   13160660
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131089"
                  Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
                  Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=131089"
                  Splits(0)._ColumnProps(11)=   "Column(1).WrapText=1"
                  Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
                  PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   0
                  ColumnFooters   =   -1  'True
                  DefColWidth     =   0
                  HeadLines       =   1.75
                  FootLines       =   1.75
                  PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
                  PictureCurrentRow(0)=   "bHQAAA4DAABCTQ4DAAAAAAAANgAAACgAAAARAAAADgAAAAEAGAAAAAAA2AIAAAAAAAAAAAAAAAAA"
                  PictureCurrentRow(1)=   "AAAAAACltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUA"
                  PictureCurrentRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAKW2taW2"
                  PictureCurrentRow(3)=   "taW2taW2taW2taW2taW2taW2taW2taW2tZSmpaW2taW2taW2taW2taW2taW2tQCtvr2tvr2tvr2t"
                  PictureCurrentRow(4)=   "vr2tvr2tvr2tvr2tvr2tvr2tvr0YGBicrq2tvr2tvr2tvr2tvr2tvr0ArcfGrcfGrcfGrcfGrcfG"
                  PictureCurrentRow(5)=   "rcfGrcfGrcfGrcfGrcfGAAAAGBgYnK6trcfGrcfGrcfGrcfGALXHxrXHxrXHxrXHxrXHxrXHxrXH"
                  PictureCurrentRow(6)=   "xrXHxrXHxrXHxgAAAAAAABgYGKW2tbXHxrXHxrXHxgC9z869z869z869z869z869z869z869z869"
                  PictureCurrentRow(7)=   "z869z84AAAAAAAAAAAAYGBitvrW9z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
                  PictureCurrentRow(8)=   "AAAAAAAAAAAAAAAAKTAxvc/Ovc/OAMbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1gAAAAAA"
                  PictureCurrentRow(9)=   "AAAAABAQEKW2tcbX1sbX1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YAAAAAAAAQEBC1"
                  PictureCurrentRow(10)=   "trXO19bO19bO19YA1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnAAAAEBAQtb7G1ufn1ufn"
                  PictureCurrentRow(11)=   "1ufn1ufnAN7n797n797n797n797n797n797n797n797n797n7xAQEL3Hxt7n797n797n797n797n"
                  PictureCurrentRow(12)=   "7wDe7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7++9z87e7+/e7+/e7+/e7+/e7+/e7+8A5+/3"
                  PictureCurrentRow(13)=   "5+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/3AA=="
                  PictureCurrentRow.vt=   9
                  PictureStandardRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
                  PictureStandardRow(0)=   "bHQAAN4DAABCTd4DAAAAAAAANgAAACgAAAARAAAAEgAAAAEAGAAAAAAAqAMAAAAAAAAAAAAAAAAA"
                  PictureStandardRow(1)=   "AAAAAACMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpwA"
                  PictureStandardRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAJyurZyu"
                  PictureStandardRow(3)=   "rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurQCltrWltrWltrWl"
                  PictureStandardRow(4)=   "trWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUArb69rb69rb69rb69rb69"
                  PictureStandardRow(5)=   "rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69AK2+va2+va2+va2+va2+va2+va2+"
                  PictureStandardRow(6)=   "va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+vQC1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1"
                  PictureStandardRow(7)=   "x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8YAtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfG"
                  PictureStandardRow(8)=   "tcfGtcfGtcfGtcfGtcfGtcfGtcfGAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
                  PictureStandardRow(9)=   "zr3Pzr3Pzr3Pzr3Pzr3PzgC9z869z869z869z869z869z869z869z869z869z869z869z869z869"
                  PictureStandardRow(10)=   "z869z869z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
                  PictureStandardRow(11)=   "vc/Ovc/OAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
                  PictureStandardRow(12)=   "1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YA1ufn"
                  PictureStandardRow(13)=   "1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnANbn59bn59bn"
                  PictureStandardRow(14)=   "59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wDe7+/e7+/e7+/e7+/e"
                  PictureStandardRow(15)=   "7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+8A3u/v3u/v3u/v3u/v3u/v3u/v"
                  PictureStandardRow(16)=   "3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/vAN7v797v797v797v797v797v797v797v"
                  PictureStandardRow(17)=   "797v797v797v797v797v797v797v797v797v7wA="
                  PictureStandardRow.vt=   9
                  MultipleLines   =   0
                  CellTipsWidth   =   0
                  MultiSelect     =   2
                  DeadAreaBackColor=   16252927
                  RowDividerColor =   13160660
                  RowSubDividerColor=   13160660
                  DirectionAfterEnter=   1
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(5)   =   ":id=0,.fontname=宋体"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF7FFFF&,.wraptext=-1,.bold=0"
                  _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(8)   =   ":id=1,.fontname=宋体"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bold=0,.fontsize=900"
                  _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.alignment=2,.valignment=2"
                  _StyleDefs(14)  =   ":id=3,.bgpicMode=2,.bold=0,.fontsize=900,.italic=0,.underline=0,.strikethrough=0"
                  _StyleDefs(15)  =   ":id=3,.charset=134"
                  _StyleDefs(16)  =   ":id=3,.fontname=宋体"
                  _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H31CFFF&"
                  _StyleDefs(19)  =   ":id=6,.fgcolor=&H80000008&"
                  _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgpicMode=2,.appearance=1"
                  _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                  _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgpicMode=2"
                  _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                  _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
                  _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(46)  =   "Named:id=33:Normal"
                  _StyleDefs(47)  =   ":id=33,.parent=0"
                  _StyleDefs(48)  =   "Named:id=34:Heading"
                  _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(50)  =   ":id=34,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
                  _StyleDefs(51)  =   "Named:id=35:Footing"
                  _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(53)  =   ":id=35,.bgpicMode=1,.bgbmp=2"
                  _StyleDefs(54)  =   "Named:id=36:Selected"
                  _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(56)  =   "Named:id=37:Caption"
                  _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(58)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(60)  =   "Named:id=39:EvenRow"
                  _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(62)  =   "Named:id=40:OddRow"
                  _StyleDefs(63)  =   ":id=40,.parent=33"
                  _StyleDefs(64)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(65)  =   ":id=41,.parent=34,.bgcolor=&HCEDFDE&,.bgpicMode=0,.borderColor=&H80000005&"
                  _StyleDefs(66)  =   "Named:id=42:FilterBar"
                  _StyleDefs(67)  =   ":id=42,.parent=33"
                  _StyleDefs(68)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(69)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(70)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(71)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(72)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(73)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(74)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(75)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(76)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(77)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(78)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(79)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(80)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(81)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(82)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(83)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(84)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(85)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(86)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(87)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(88)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(89)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(90)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(91)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(92)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(93)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(94)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(95)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
                  _StyleDefs(96)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(97)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(98)  =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(99)  =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(100) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(101) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(102) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(103) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(104) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(105) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(106) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(107) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(108) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(109) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(110) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(111) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(112) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(113) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(114) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(115) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(116) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(117) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(118) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(119) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(120) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(121) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(122) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(123) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               End
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
               Height          =   3855
               Left            =   11610
               TabIndex        =   15
               Top             =   510
               Width           =   10965
               _ExtentX        =   19341
               _ExtentY        =   6800
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
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mvarObjectID As String


Private thsRs01 As New RecordSet  '第一个标签页的记录集
Private thsRs02 As New RecordSet  '第二个标签页的记录集
Private strSQL As String

Private theStartDate As String
Private theEndDate As String    '起止时间
Private theOrderCode As String   '订单号
Private theClient As String   '客户

Public Property Let ObjectID(ByVal vData As String)
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = mvarObjectID
End Property

Private Sub GetQueryParameters()
'    theStartDate = DateTimePicker1.Value
'    theEndDate = DateTimePicker2.Value
    theStartDate = Format(DTPicker1.Value, "YYYY-MM-DD")
    theEndDate = Format(DTPicker2.Value, "YYYY-MM-DD")
    theOrderCode = Text1.Text
'    theClient = FlatEdit3.Text
End Sub

Private Sub InitLayout()
    With ActiveBar21
        .ClientAreaControl = C1Elastic2
        .RecalcLayout
    End With
    
    InitDatetime
End Sub

Private Sub InitFrm()
theClient = ""
    InitLayout
    GetData
End Sub

Private Sub InitDatetime()
    Dim szStartDate As String
    Dim szEndDate As String
    
    szEndDate = Format(Now, "YYYY-MM-DD")
    szStartDate = DateAdd("m", -1, szEndDate)
    
    DTPicker1.Value = szStartDate
    DTPicker2.Value = szEndDate
    
    
'    DateTimePicker1.Value = szStartDate
'    DateTimePicker2.Value = szEndDate
End Sub

Private Sub GetData()
    '获取参数
    GetQueryParameters
    
    
    
    '根据参数获取结果数据集
    GetData01
    GetData02
End Sub

'第一个标签页的网格获取并显示数据
Private Sub GetData01()
    Dim rs As New RecordSet
    Dim strSQL As String
    strSQL = "exec dbo.usp_schedule '" & theStartDate & "','" & theEndDate & "','" & theOrderCode & "','" & theClient & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    Debug.Print strSQL
    FillUnConnectRecordSet rs, thsRs01
    
    rs.Close
    Set rs = Nothing
    
    TDBGrid1.DataSource = thsRs01
    
    '设置网格样式
    FormatGrid01
End Sub

'第二个标签页的网格获取并显示数据
Private Sub GetData02()
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim a As String
    Dim b As String
    
    a = Format(DTPicker1.Value, "YYYY-MM-DD")
    b = Format(DTPicker2.Value, "YYYY-MM-DD")
    strSQL = "exec usp_schedule_2 '" & a & "','" & b & "','" & Text1.Text & "','" & theClient & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    FillUnConnectRecordSet rs, thsRs02
    rs.Close
    Set rs = Nothing
    
    TDBGrid2.DataSource = thsRs02
    
    '设置网格样式
    FormatGrid02

End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "查询"
        GetData
        Case "退出"
            Unload Me
        Case "保存"
            
        Case "设置数量"
            
            
    End Select
End Sub

Private Sub Form_Load()
    InitFrm
    
    
    
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


Private Sub FormatGrid01()
    TDBGrid1.Columns("B_DepartFb").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_DepartDJu").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_Qprocess").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_Hprocess").ValueItems.Presentation = dbgCheckBox
    
    
    TDBGrid1.Columns("B_pactcode").Locked = True
    TDBGrid1.Columns("B_ClientName").Locked = True
    TDBGrid1.Columns("B_Date").Locked = True
    TDBGrid1.Columns("B_YarnInsert").Locked = True
    TDBGrid1.Columns("B_ProcessInsert").Locked = True
    TDBGrid1.Columns("B_ProcessDelivery").Locked = True
    
    TDBGrid1.Columns("B_pactcode").Caption = "订单号"
    TDBGrid1.Columns("B_ClientName").Caption = "染厂"
    TDBGrid1.Columns("B_Date").Caption = "订单日期"
    TDBGrid1.Columns("B_YarnInsert").Caption = "原料入库"
    TDBGrid1.Columns("B_YarnDelivery").Caption = "原料发货"
    TDBGrid1.Columns("B_WhiteInsert").Caption = "白坯入库"
    TDBGrid1.Columns("B_WhiteDelivery").Caption = "白坯发货"
'    TDBGrid1.Columns("B_DepartFb").Caption = "打样/制版"
'    TDBGrid1.Columns("B_DepartColor").Caption = "计划"

'    TDBGrid1.Columns("B_DepartFb").Locked = False
'    TDBGrid1.Columns("B_DepartColor").Locked = False
'    TDBGrid1.Columns("B_DepartDJu").Locked = False
    TDBGrid1.Columns("B_DepartColor").Visible = False
    TDBGrid1.Columns("B_DepartColor").Locked = True
    TDBGrid1.Columns("B_DepartColor").AllowSizing = False
    TDBGrid1.Columns("B_DepartFb").Visible = False
    TDBGrid1.Columns("B_DepartFb").Locked = True
    TDBGrid1.Columns("B_DepartFb").AllowSizing = False
    
    TDBGrid1.Columns("B_DepartDJu").Caption = "染色/印花"
    TDBGrid1.Columns("B_ProcessInsert").Caption = "深加工入库"
    TDBGrid1.Columns("B_ProcessDelivery").Caption = "深加工发货"
    TDBGrid1.Columns("B_Qprocess").Caption = "白坯前处理"
    TDBGrid1.Columns("B_Hprocess").Caption = "后整理"

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
'    TDBGrid1.HoldFields
    
    TDBGrid1.MarqueeStyle = dbgSolidCellBorder
End Sub
Private Sub FormatGrid02()
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
    TDBGrid2.Columns("B_ClientName").Caption = "染厂"
    TDBGrid2.Columns("B_Date").Caption = "订单日期"
    TDBGrid2.Columns("B_YarnInsert").Caption = "原料入库"
    TDBGrid2.Columns("B_YarnDelivery").Caption = "原料发货"
    TDBGrid2.Columns("B_WhiteInsert").Caption = "白坯入库"
    TDBGrid2.Columns("B_WhiteDelivery").Caption = "白坯发货"
'    TDBGrid2.Columns("B_DepartFb").Caption = "打样/制版"
'    TDBGrid2.Columns("B_DepartColor").Caption = "计划"
    TDBGrid2.Columns("B_DepartDJu").Caption = "染色/印花"
    TDBGrid2.Columns("B_ProcessInsert").Caption = "深加工入库"
    TDBGrid2.Columns("B_ProcessDelivery").Caption = "深加工发货"
    TDBGrid2.Columns("B_Qprocess").Caption = "白坯前处理"
    TDBGrid2.Columns("B_Hprocess").Caption = "后整理"
    
    TDBGrid2.Columns("B_DepartFb").Visible = False
    TDBGrid2.Columns("B_DepartFb").Locked = True
    TDBGrid2.Columns("B_DepartFb").AllowSizing = False
    TDBGrid2.Columns("B_DepartColor").Visible = False
    TDBGrid2.Columns("B_DepartColor").Locked = True
    TDBGrid2.Columns("B_DepartColor").AllowSizing = False
    
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
    
    
End Sub

Private Sub bianse()
    Dim i As Long
    For i = 0 To TDBGrid2.Columns.Count - 1
            TDBGrid2.Columns(i).FetchStyle = True
    Next
End Sub

Private Sub FillUnConnectRecordSet(ByRef sRs As RecordSet, ByRef tRs As RecordSet)
    On Error Resume Next
    Dim i As Long
       
    Set tRs = New RecordSet
    For i = 0 To sRs.Fields.Count - 1
        tRs.Fields.Append sRs.Fields(i).name, sRs.Fields(i).Type, sRs.Fields(i).DefinedSize, sRs.Fields(i).Attributes
    Next
       
    tRs.Open
    Do While Not sRs.EOF
        tRs.AddNew
        For i = 0 To sRs.Fields.Count - 1
            tRs.Fields(Trim(sRs(i).name)).Value = IIf(IsNull(sRs.Fields(Trim(sRs.Fields(i).name)).Value), Null, sRs.Fields(Trim(sRs.Fields(i).name)).Value)
        Next
        tRs.Update
        sRs.movenext
    Loop
      
    tRs.MoveFirst
End Sub

Private Sub PushButton1_Click()
    Dim frm1 As New frmPopupDanWei
    frm1.ContactType = "染厂"
    frm1.Show vbModal
'    Originalsuppliers = frm1.clientid
     theClient = frm1.clientid   '全局变量   接收客户ID
    FlatEdit3.Text = frm1.ClientName
    Unload frm1
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

    If Len(IIf(IsNull(thsRs01!B_id), "", thsRs01!B_id)) <= 0 Then
        sql = "insert into G_schedule (B_orderid," & b & ") values('" & thsRs01!B_Pactid & "','" & a & "')"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
 
    Else
        sql = "update G_schedule set " & b & "='" & a & "' where B_id='" & thsRs01!B_id & "'"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql

    End If

   GetData
    TDBGrid1.bookmark = bookmark
End Sub

Private Sub TDBGrid1_Click()
    If TDBGrid1.Col <= 13 And TDBGrid1.Col > 10 Then
                'setlogo
                Exit Sub
    End If
End Sub

Private Sub TDBGrid1_ColEdit(ByVal colIndex As Integer)
    Dim rowIndex As Long
    rowIndex = TDBGrid1.bookmark
    
    
    If TDBGrid1.Col <= 13 And TDBGrid1.Col > 10 Then
        'MsgBox "更新DB，行=" & rowIndex & "，列=" & colIndex
        setlogo
    End If
End Sub

Private Sub TDBGrid2_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, _
    bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    Dim i As Long
    Dim j As Long
    Dim m_Num As Long
    '需要做的工序 -已刷卡
        If Val(TDBGrid2.Columns(Col).CellValue(bookmark)) > 0 And Val(TDBGrid2.Columns(Col).CellValue(bookmark)) < 1 Then
            If Col > 4 Then
                CellStyle.BackColor = vbGreen
                CellStyle.ForeColor = CellStyle.BackColor
            End If

        End If
    '需要做的工序 - 未刷卡
    
    If Val(TDBGrid2.Columns(Col).CellValue(bookmark)) = 0 Then

        If Col > 4 Then
            CellStyle.BackColor = vbRed
            
            CellStyle.ForeColor = CellStyle.BackColor
        End If

    End If

    
    Debug.Print Val(TDBGrid2.Columns(Col).CellValue(bookmark))
    '不需要做的工序
    If Val(TDBGrid2.Columns(Col).CellValue(bookmark)) >= 1 Then
        If Col > 4 Then
            CellStyle.BackColor = &HFFFF&
            
            CellStyle.ForeColor = CellStyle.BackColor
        End If
    End If
    
End Sub

Private Sub setlogo()
Dim sql As String
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long
    
    i1 = Abs(TDBGrid1.Columns("B_Qprocess").Value)
    i2 = Abs(TDBGrid1.Columns("B_DepartDJu").Value)
    i3 = Abs(TDBGrid1.Columns("B_Hprocess").Value)
    
        If Len(IIf(IsNull(thsRs01!B_id), "", thsRs01!B_id)) <= 0 Then
        sql = "insert into G_schedule (B_orderid,B_Qprocess,B_DepartDJu,B_Hprocess) values"
         sql = sql & "('" & thsRs01!B_Pactid & "','" & i1 & "','" & i2 & "','" & i3 & "')"
        'sql = "insert into G_schedule (B_orderid,B_DepartFb,B_DepartDJu) values('" & thsRs01!B_Pactid & "','" & m & "','" & n & "')"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
    Else
        sql = "update G_schedule SET B_Qprocess='" & i1 & "',B_DepartDJu='" & i2 & "',B_Hprocess='" & i3 & "'  "
        sql = sql & " where B_id='" & thsRs01!B_id & "'"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
    End If
End Sub
