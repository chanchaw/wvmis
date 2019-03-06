VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmModBLSource 
   Caption         =   "原料入库"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11685
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
   ScaleHeight     =   7020
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Tag             =   "12B003"
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7020
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11685
      _cx             =   20611
      _cy             =   12383
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
      BackColor       =   13557726
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
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
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmModBLSource.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.PictureBox PictureButton 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   6210
         ScaleHeight     =   690
         ScaleWidth      =   5475
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   6330
         Width           =   5475
         Begin XtremeSuiteControls.PushButton UCButton1 
            Height          =   435
            Index           =   0
            Left            =   1800
            TabIndex        =   5
            Top             =   120
            Width           =   1155
            _Version        =   1048578
            _ExtentX        =   2037
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "保存"
            BackColor       =   13557726
            UseVisualStyle  =   -1  'True
            ImageGap        =   11
            IconWidth       =   16
            Icon            =   "frmModBLSource.frx":0058
         End
         Begin XtremeSuiteControls.PushButton UCButton1 
            Height          =   435
            Index           =   1
            Left            =   3000
            TabIndex        =   6
            Top             =   120
            Width           =   1155
            _Version        =   1048578
            _ExtentX        =   2037
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "保存草稿"
            BackColor       =   13557726
            UseVisualStyle  =   -1  'True
            IconWidth       =   16
            Icon            =   "frmModBLSource.frx":04C2
         End
         Begin XtremeSuiteControls.PushButton UCButton1 
            Height          =   435
            Index           =   2
            Left            =   4200
            TabIndex        =   7
            Top             =   120
            Width           =   1155
            _Version        =   1048578
            _ExtentX        =   2037
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "关闭"
            BackColor       =   13557726
            UseVisualStyle  =   -1  'True
            ImageGap        =   11
            IconWidth       =   16
            Icon            =   "frmModBLSource.frx":092C
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00CEDFDE&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   0
         ScaleHeight     =   690
         ScaleWidth      =   6210
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   6330
         Width           =   6210
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   3435
         End
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 ABMenu 
         Height          =   750
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11685
         _LayoutVersion  =   1
         _ExtentX        =   20611
         _ExtentY        =   1323
         _DataPath       =   ""
         Bands           =   "frmModBLSource.frx":0D96
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   7740
            Top             =   60
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   2520
            Top             =   1800
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Timer Timer1 
            Left            =   5580
            Top             =   120
         End
      End
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
         Height          =   5580
         Left            =   0
         TabIndex        =   8
         Top             =   750
         Width           =   11685
         _LayoutVersion  =   1
         _ExtentX        =   20611
         _ExtentY        =   9843
         _DataPath       =   ""
         Bands           =   "frmModBLSource.frx":CDC8
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Bindings        =   "frmModBLSource.frx":CF90
            Height          =   4515
            Left            =   600
            TabIndex        =   12
            Top             =   240
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   7964
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
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
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
         Begin VB.PictureBox PctBack 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            Begin VB.CommandButton btPopUpWindow 
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Marlett"
                  Size            =   9
                  Charset         =   2
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.TextBox txtNullTip 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   555
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "无数据，请使用右键“增加明细”"
            Top             =   2220
            Width           =   6675
         End
      End
   End
End
Attribute VB_Name = "frmModBLSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mvarObjectID As String
Public m_ID As Long
Private m_ReportObjectID As String
Private WithEvents clsBL1 As clsBL
Attribute clsBL1.VB_VarHelpID = -1
'Private WithEvents clsBL1 As clsBLOri
Private clsCtlShow1 As clsCtlShow
Private WithEvents clsGridShow1 As clsGridShow
Attribute clsGridShow1.VB_VarHelpID = -1


Private g_OptionButtonNumber As Long
Private g_BoolColIndex As Boolean   'true为有列序号,false为没有

Private m_TaxRate As Double
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private WithEvents A_clsKeyDetec As clsKeyDetec
Attribute A_clsKeyDetec.VB_VarHelpID = -1
Private clsDataType1 As New clsDataType

Public Property Let ObjectID(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ObjectID
    ObjectID = mvarObjectID
End Property

Private Sub A_clsKeyDetec_AfterTimes()
    CopyDetailOne Adodc2.RecordSet!B_itemid
End Sub



Private Sub clsBL1_OnAddNewItem(Adodc2 As Object, Cancel As Integer)
    If m_TaxRate <> 0 Then
        Adodc2.RecordSet("B_TaxRate") = m_TaxRate
    End If
End Sub

Private Sub clsBL1_OnAddNewObject(Adodc1 As Object)
    ClearListBoxContent
    clsCtlShow1.LoadObject Adodc1.RecordSet
    
    AutoFillCheckBy
End Sub

Private Sub clsBL1_OnInitFrame()
    Set clsGridShow1 = New clsGridShow
    With clsGridShow1
        Set .fatherFrm = Me
        .ObjectID = mvarObjectID
        .InitClass TDBGrid1, 4
        .ShowGridFormat
        .ShowGridCtl
    End With
End Sub

Private Sub clsBL1_OnOpenFrame()
    ClearListBoxContent
    '当打开单据时
    clsCtlShow1.LoadObject Adodc1.RecordSet
    
    SetNullTip
End Sub

Private Sub clsBL1_OnSaveFrame()
    clsCtlShow1.SaveObject Adodc1.RecordSet
    '当保存单据时
    TDBGrid1.Update
    clsBL1.boolIsSave = False
End Sub

Private Sub clsBL1_OnUpdateFrameType()
    '当 系统菜单项需要修改时进行
End Sub

Private Sub Form_Load()
'    prevWndProc = GetWindowLong(TDBGrid1.hWnd, GWL_WNDPROC)
'    SetWindowLong TDBGrid1.hWnd, GWL_WNDPROC, AddressOf WndProc
    GetObjectParameter
    With ActiveBar21
        .ClientAreaControl = TDBGrid1
        .RecalcLayout
    End With
    AnimateForm Me
    InitClass
    
    ActiveBar21.RecalcLayout
    
    SetNullTip
End Sub

Private Sub InitClass()
    Set clsCtlShow1 = New clsCtlShow
    With clsCtlShow1
        .ObjectID = mvarObjectID
        .InitClass ActiveBar21, 2
        .Refresh
    End With
    
    Set clsBL1 = New clsBL
    'Set clsBL1 = New clsBLOri
    With clsBL1
        .ObjectID = mvarObjectID
        .InitClass Adodc1, Adodc2, TDBGrid1, ABMenu
    End With


    '初始化复制数据用到的快捷键
    Set A_clsKeyDetec = New clsKeyDetec
    A_clsKeyDetec.InitCls Me, 46, 2, Timer1  '点击两次DEL按键触发复制操作
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ABMenu.Bands("P操作").PopupMenu
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If clsBL1.ExitFrame = False Then

        Cancel = 1
    Else
        clsCtlShow1.RemoveAll

    End If
    
    Gm.CacheFrms.DelFrm mvarObjectID
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    Dim m_Mark As Integer
    Select Case ColIndex
        Case 0
            m_Mark = Adodc2.RecordSet.Bookmark
            Adodc2.RecordSet.requery
            Adodc2.RecordSet.Bookmark = Val(m_Mark)
    End Select
End Sub

Private Sub TDBGrid1_AfterDelete()
    SetNullTip
End Sub

Private Sub TDBGrid1_AfterInsert()
    SetNullTip
End Sub

Private Sub TDBGrid1_ButtonClick(ByVal ColIndex As Integer)

    If TDBGrid1.Columns("B_OrderCode").ColIndex = ColIndex Then
'
    Dim frm1 As New frmPopupDingdanhao
    frm1.Show vbModal
    Dim a As String
    a = frm1.departmentid
    If a <> "" Then
    Adodc2.RecordSet("b_OrderCode") = a
    End If
    Unload frm1

    End If
End Sub

Private Sub TDBGrid1_Change()
    SetNullTip
End Sub

Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
'        TDBGrid1.Update
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And clsBL1.intIsClosed = 0 Then
        ABMenu.Bands("P明细").PopupMenu
    End If
End Sub

Private Sub UCButton1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            Save
        Case 1
            clsBL1.SaveFrame
        Case 2
            Unload Me
    End Select
End Sub




Public Sub LoadObject()
    clsBL1.LoadObject
    
    TDBGrid1.Columns("B_OrderCode").Button = True
'    TDBGrid1.Columns("B_GoodsID").Button = True
End Sub

'新增对象
Public Sub AddNewObject()
    clsBL1.AddNewFrame
End Sub

'编辑对象
Public Sub EditObject(ByVal m_KeyID As Variant)
    On Error Resume Next
    m_ID = m_KeyID
    clsBL1.m_ID = m_KeyID
    clsBL1.OpenFrame
End Sub

'取得参数
Private Sub GetObjectParameter()
    On Error Resume Next
    Dim rs As New RecordSet
    Dim strSQL As String

    Set rs = New RecordSet
    strSQL = "Select B_Width,B_Height,B_BillName From G_BL Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.width = rs("B_Width")
    Me.height = rs("B_Height")
    Me.Caption = rs("B_BillName")
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub Change(ByVal sCtl As String, ByVal sCommand As String)
    If sCtl = "B_InvoiceType" Then
        m_TaxRate = Me.Controls("B_InvoiceType").ReturnValue(1)
        If Adodc2.RecordSet.RecordCount > 0 Then
            If clsBL1.intIsClosed <> 1 Then
                SetTaxRate
            End If
        End If
    End If
End Sub

Private Sub SetTaxRate()
    On Error Resume Next
    Dim m_Mark
    
    If Adodc2.RecordSet.RecordCount > 0 Then
        m_Mark = Adodc2.RecordSet.Bookmark
        Adodc2.RecordSet.MoveFirst
    Else
        Exit Sub
    End If
    
    Do While Not Adodc2.RecordSet.EOF
        Adodc2.RecordSet("B_TaxRate") = m_TaxRate
        Adodc2.RecordSet.Update
        
        If Adodc2.RecordSet("B_Qty") <> 0 Then
            clsBL1.UseFormulaCount
        End If
        Adodc2.RecordSet.movenext
    Loop
    
    Adodc2.RecordSet.Bookmark = m_Mark
    clsBL1.UpdateSum
End Sub

Private Sub AutoFillCheckBy()
    On Error Resume Next
    If clsBL1.boolIsDraft = True Then
        Me.Controls("B_UserName").Text = Gm.SysID.SystemUser
    End If
End Sub


Private Sub ClearListBoxContent()
    On Error Resume Next
    Dim oUCListBox As Object
    For Each oUCListBox In Me.Controls
        If TypeName(oUCListBox) = "UCListBox" Then
            oUCListBox.Text = ""
        End If
    Next
    
    Me.Controls("B_UserName").Text = ""
End Sub



'草稿状态下复制当前所在行的明细数据
Private Sub CopyDetailOne(ByVal vItemID As Long)
    Dim szFields As String
    Dim cls1 As New clsDataBase
    Dim m_DraftDetailTable As String
    
    
    m_DraftDetailTable = clsBL1.DraftDetailTable
    szFields = cls1.GetTableFields(m_DraftDetailTable)
    Debug.Print szFields
    
    strSQL = "INSERT INTO " & m_DraftDetailTable
    strSQL = strSQL & " (" & szFields & ")"
    strSQL = strSQL & " Select " & szFields & " From " & m_DraftDetailTable & " WHERE B_ItemID='" & vItemID & "'"
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
    
    Adodc2.RecordSet.requery
    Adodc2.RecordSet.movelast
End Sub


Private Sub SetNullTip()
    If TDBGrid1.ApproxCount <= 0 Then
        txtNullTip.Visible = True
    Else
        txtNullTip.Visible = False
    End If
End Sub


Private Sub SaveStyle()
    
End Sub


Private Sub ABMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.Name
    
        Case "保存样式"
            clsGridShow1.SaveColWidth
    End Select
End Sub



'本代码块是检测数值列仅可输入数字
'=============================================
Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, _
    ByVal KeyAscii As Integer, Cancel As Integer)
    
    On Error Resume Next
    Cancel = 0
    
    If IsNumericLegal(ColIndex, KeyAscii) = False Then
        MsgBox "只可以输入数字！", vbOKOnly + vbInformation, "提示"
        Cancel = 1
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, _
    OldValue As Variant, Cancel As Integer)
    
    Dim szContent As String
    
    If clsDataType1.IsNumeric(Adodc2.RecordSet.Fields(ColIndex)) = True Then
        szContent = TDBGrid1.Columns(ColIndex).Text
        If Len(CStr(szContent)) <= 0 Then
            Cancel = 1
        End If
    End If
End Sub


Private Function IsNumericLegal(ByVal ColIndex As Integer, _
    ByVal KeyAscii As Integer) As Boolean
    
    IsNumericLegal = True
    
    If clsDataType1.IsNumeric(Adodc2.RecordSet.Fields(ColIndex)) = True Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii = 8 Or KeyAscii = 46 Then
                IsNumericLegal = True
            Else
                IsNumericLegal = False
            End If
            
        End If
    End If
End Function
'=============================================

'保存单据
Private Sub Save()
    On Error Resume Next
    If JudgeBeforeSave = False Then
        Exit Sub
    End If
    
    clsBL1.BillCheckIn
End Sub

'返回FALSE表示不可保存，相反为可以保存
Private Function JudgeBeforeSave() As Boolean
    '默认为TRUE，可以保存
    JudgeBeforeSave = True
    
    '1. 判断单据日期和当前日期是否相同
    JudgeBeforeSave = JudgeBillDate
    If JudgeBeforeSave = False Then
        Exit Function
    End If
    
    '2. 开发环境中关于主表字段的非法判断
    JudgeBeforeSave = clsCtlShow1.JudgeBeforeSave
    If JudgeBeforeSave = False Then
        Exit Function
    End If
End Function

Private Function JudgeBillDate() As Boolean
    Dim szDate As String
    szDate = Me.Controls("B_Date").Text
    JudgeBillDate = IsToday(szDate)
End Function


