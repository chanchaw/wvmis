Attribute VB_Name = "ABUtilities"
Option Explicit

'制作左侧导航栏的辅助工具类
Public Sub ABAddFlag(ByVal bandFlag As ActiveBar2LibraryCtl.BandFlags, ByVal band As ActiveBar2LibraryCtl.band)
    band.flags = band.flags Or bandFlag
End Sub

Public Sub ABRemoveFlag(ByVal bandFlag As ActiveBar2LibraryCtl.BandFlags, ByVal band As ActiveBar2LibraryCtl.band)
    band.flags = band.flags And Not bandFlag
End Sub


Public Function GetUniqueToolID() As Long
Static STATToolId As Long

If STATToolId = 0 Then
    STATToolId = 20000
End If

STATToolId = STATToolId + 1

GetUniqueToolID = STATToolId

End Function


Public Function getSeparator(ByVal ActiveBarTarget As ActiveBar2LibraryCtl.ActiveBar2) As ActiveBar2LibraryCtl.Tool
Dim o As Object

On Error Resume Next
    
    Set o = ActiveBarTarget.Tools("sep")
    
    If o Is Nothing Then
        Set o = ActiveBarTarget.Tools.Add(GetUniqueToolID(), "sep")   ' should we still have Tool.BeginGroup ? **************
        o.ControlType = ddTTSeparator
        o.TagVariant = "sep"
    End If
    
    Set getSeparator = o
    
End Function



