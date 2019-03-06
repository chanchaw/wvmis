VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmNavigatorAccessory 
   Caption         =   "辅料模块"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNavigatorAccessory.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10380
      _LayoutVersion  =   1
      _ExtentX        =   18309
      _ExtentY        =   12991
      _DataPath       =   ""
      Bands           =   "frmNavigatorAccessory.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5895
         Left            =   540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   9615
         _cx             =   16960
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
         _GridInfo       =   $"frmNavigatorAccessory.frx":0752
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   5835
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   9555
            _cx             =   16854
            _cy             =   10292
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
            Appearance      =   2
            MousePointer    =   0
            Version         =   800
            BackColor       =   16777215
            ForeColor       =   -2147483630
            FrontTabColor   =   14270310
            BackTabColor    =   16777215
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "Tab&1|Tab&2|Tab&3"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   6
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
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
            Flags(0)        =   2
            Flags(1)        =   2
            Flags(2)        =   2
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   5250
               Left            =   45
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   540
               Width           =   9465
               _cx             =   16695
               _cy             =   9260
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
               BorderWidth     =   1
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
               _GridInfo       =   $"frmNavigatorAccessory.frx":07D2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   5220
                  Left            =   15
                  ScaleHeight     =   5220
                  ScaleWidth      =   9435
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   9435
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   0
                     Left            =   900
                     TabIndex        =   5
                     Tag             =   "19B007"
                     Top             =   300
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "辅料出库"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorAccessory.frx":084E
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   1
                     Left            =   900
                     TabIndex        =   6
                     Tag             =   "13B002"
                     Top             =   2340
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "单据流水"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorAccessory.frx":2CB8
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   2
                     Left            =   3600
                     TabIndex        =   7
                     Tag             =   "13B003"
                     Top             =   2340
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "库存表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorAccessory.frx":5122
                  End
                  Begin XtremeSuiteControls.PushButton btnObject 
                     Height          =   1335
                     Index           =   3
                     Left            =   6300
                     TabIndex        =   8
                     Tag             =   "13B004"
                     Top             =   2340
                     Width           =   1395
                     _Version        =   1048578
                     _ExtentX        =   2461
                     _ExtentY        =   2355
                     _StockProps     =   79
                     Caption         =   "收发存汇总表"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     UseVisualStyle  =   -1  'True
                     TextImageRelation=   1
                     IconWidth       =   48
                     Icon            =   "frmNavigatorAccessory.frx":758C
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmNavigatorAccessory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitLayout()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub

Public Sub InitFrm()
    InitLayout
    
    ConfirmPermission
End Sub

Private Sub btnObject_Click(Index As Integer)
    OpenObject btnObject(Index).Tag, "通过导航中按钮打开"
End Sub

Private Sub Form_Load()
    InitFrm
End Sub


Private Sub OpenObject(ByVal m_ObjectID As String, ByVal m_BillName As String)
    Gm.Authority.Execute m_ObjectID, m_BillName, "LoadObject", Nothing
End Sub



'设置按钮的可用度
Public Sub ConfirmPermission()
    On Error Resume Next
    Dim i As Long
    Dim szObjectID As String
    
    
    For i = 0 To btnObject.Count - 1
        szObjectID = btnObject(i).Tag
        If Len(szObjectID) > 0 Then
            If Gm.PI.JudgeView(szObjectID) = True Then
                btnObject(i).Enabled = True
            Else
                btnObject(i).Enabled = False
            End If
        End If
    Next
End Sub

