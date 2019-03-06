Attribute VB_Name = "modAnimateForm"
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type RECT

    Left As Long '矩形左上角的X坐标
    Top As Long '`矩形左上角的Y坐标
    Right As Long '`矩形右下角的X坐标
    Bottom As Long '`矩形右下角的Y坐标
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Const RGN_DIFF = 4
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80

Dim LastLeft  As Single
Dim LastTop  As Single

Public frmMe
Dim a As RECT
Public Picture1 As Object

Public Sub AnimateForm(ByRef frm As Object)
    If Gm.iAnimate = 1 Then
        Exit Sub
    End If
    On Error Resume Next
    Set frmMe = frm
    frmMe.Visible = False
    Dim iR As Long

    iR = GetWindowRect(frmMe.hWnd, a)
    
    Set Picture1 = frmMe.Controls.Add("VB.PictureBox", "Picture9999")

    With Picture1
        .Appearance = 0
        .BorderStyle = 1
        .BackColor = &H80000003
        .ZOrder 1
    End With
    SetWindowLong Picture1.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    SetParent Picture1.hWnd, 0
    AutoMove
    
    FrmMePaint
    Picture1.Visible = False
    
    SetParent Picture1.hWnd, frmMe.hWnd

    frmMe.Controls.Remove ("Picture9999")
    Set Picture1 = Nothing
    Set frmMe = Nothing
End Sub

'从小到大
Public Sub AutoMove()
    Dim i As Integer
    Dim iSplit As Integer
    Dim xWidth As Long     'x等分
    Dim yHeight As Long     'y等分
    
    Dim iTop As Integer
    Dim iLeft As Integer

    Dim MouseXY As POINTAPI
    Dim lReturn As Long
    
    Dim lsLeft      As Long
    Dim lsTop       As Long
    
    Dim Xmove As Single
    Dim Ymove As Single
    
    Dim tmpleft As Single
    Dim tmpTop As Single
    
    Dim MeH As Single
    Dim MeW As Single
    
    iSplit = 10

    xWidth = frmMe.Width / iSplit
    yHeight = frmMe.Height / iSplit
    
    '----起始动画最终的位置
    iTop = (Screen.Height + frmMe.Height) / 2 - yHeight
    
    iLeft = (Screen.Width - xWidth) / 2
    
    '取得鼠标坐标
    lReturn = GetCursorPos(MouseXY)
    
    '----取得起始动画起始座标
    
    lsLeft = frmMe.ScaleX(MouseXY.X, vbPixels, vbTwips)
    lsTop = frmMe.ScaleY(MouseXY.Y, vbPixels, vbTwips)
    
    '----移动的步进大小
    Xmove = (iLeft - lsLeft) / iSplit
    Ymove = (iTop - lsTop) / iSplit
    
    '----起始到终止动画
    For i = 1 To iSplit
        With Picture1
            .Visible = False
            .Move lsLeft + i * Xmove, lsTop + i * Ymove, xWidth, yHeight
        End With

        ShowPicture
        
        If i > 6 Then
            Sleep (5)
        Else
            Sleep (10)
        End If
        Picture1.Visible = True
    Next
    tmpleft = lsLeft + (i - 1) * Xmove

    tmpTop = lsTop + (i - 1) * Ymove + yHeight

    Xmove = frmMe.Width / 2 / iSplit
    Ymove = frmMe.Height / 2 / iSplit


    MeH = frmMe.Height / iSplit
    MeW = frmMe.Width / iSplit

    Sleep (150)
    For i = 1 To iSplit - 1
        With Picture1
            .Move tmpleft - i * Xmove, tmpTop - MeH * (i + 1), MeW * (i + 1), MeH * (i + 1)
        End With
    
        ShowPicture
        
        Sleep (10 * (i * 0.25))
        Picture1.Visible = True
        
    Next
    
    LastLeft = tmpleft - (i - 1) * Xmove
    LastTop = tmpTop - MeH * i
        
End Sub

Public Sub ShowPicture()
    Dim outer_rgn As Long
    Dim inner_rgn As Long
    Dim combined_rgn As Long
    Dim wid As Single
    Dim hgt As Single
    
    ' Create the regions.
    wid = Picture1.ScaleX(Picture1.Width, vbTwips, vbPixels)
    hgt = Picture1.ScaleY(Picture1.Height, vbTwips, vbPixels)
    
    '外部显示部分
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    '内部显示部分
    inner_rgn = CreateRectRgn(1, 1, wid - 1, hgt - 1)
    
    combined_rgn = CreateRectRgn(0, 0, wid, hgt)
    CombineRgn combined_rgn, outer_rgn, inner_rgn, RGN_DIFF

    SetWindowRgn Picture1.hWnd, combined_rgn, True
End Sub

'----将窗体移到中间
Public Sub FrmMePaint()
  
    On Error Resume Next
    
    If frmMe.MDIChild Then
        Dim r As Long
        Dim a As Long, b As Long
        Dim T As Long, L As Long
        Dim W As Long, H As Long

        
        Dim pa As POINTAPI
        r = ClientToScreen(Gm.frmMain.hWnd, pa)
        
        L = frmMe.ScaleX(LastLeft, vbTwips, vbPixels)
        T = frmMe.ScaleY(LastTop, vbTwips, vbPixels)
    
        W = frmMe.ScaleX(frmMe.Width, vbTwips, vbPixels)
        H = frmMe.ScaleY(frmMe.Height, vbTwips, vbPixels)
        
        a = frmMe.ScaleX(Gm.frmMain.ActiveBar21.ClientAreaLeft, vbTwips, vbPixels)
        b = frmMe.ScaleY(Gm.frmMain.ActiveBar21.ClientAreaTop, vbTwips, vbPixels)
    
        r = MoveWindow(frmMe.hWnd, L - pa.X - a, T - b - pa.Y, W, H, 0)
    
    Else
           frmMe.Move LastLeft, LastTop

    End If
    
End Sub

