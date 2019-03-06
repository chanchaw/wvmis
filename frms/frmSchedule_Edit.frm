VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmSchedule_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置数量"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedule_Edit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7125
   StartUpPosition =   2  '屏幕中心
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   840
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
      _Version        =   1048578
      _ExtentX        =   3836
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "保存"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   375
      Left            =   3180
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _Version        =   1048578
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   840
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
      _Version        =   1048578
      _ExtentX        =   3836
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "取消退出"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   1740
      TabIndex        =   3
      Top             =   540
      Width           =   1215
      _Version        =   1048578
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "数   量:"
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
End
Attribute VB_Name = "frmSchedule_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bool As Boolean

Private Sub FlatEdit2_KeyPress(KeyAscii As Integer)
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
Private Sub PushButton2_Click()
    If Val(FlatEdit2.Text) <= 0 Then
        MsgBox "数量不能为0", vbInformation, "提示"
    End If
    bool = True
    Me.Hide
    
End Sub

Private Sub PushButton3_Click()
    Unload Me
End Sub
