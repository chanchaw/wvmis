VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDock2AB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Interface implemented by Windows that can dock to an ActiveBar

''''''''''''''''
' Description:  Instructs the window to create a band and tool and dock itself to the specified ActiveBar.
' Parameters:   ActiveBar: The ActiveBar the window should dock itself to.
' Returns:      The dockable form conainer ActiveBar Band that was created by the window in order to dock.
Public Function DockYourselfTo(ByVal ActiveBar As ActiveBar2LibraryCtl.IActiveBar2, _
                                Optional ByVal parmIsVisible As Boolean = True, _
                                Optional ByVal paramDockingarea As ActiveBar2LibraryCtl.DockingAreaTypes = ddDARight, _
                                Optional ByVal paramGrabHandleStyle As ActiveBar2LibraryCtl.GrabHandleStyles = ddGSCaption, _
                                Optional ByVal paramDockingOffset As Long = 0) As ActiveBar2LibraryCtl.IBand

End Function


''''''''''''''''''''''''''''''''''''
' You should be able to just copy the below code into a form
' in order to make it easily dockable by just calling
' Form.DockYourselfTo(...)
''''''''''''''''''''''''''''''''''''
'Implements IWillDockToActiveBar
'
'Private Function IWillDockToActiveBar_DockYourselfTo(ByVal ActiveBar As ActiveBar2LibraryCtl.IActiveBar2, Optional ByVal parmIsVisible As Boolean = True, Optional ByVal paramDockingarea As ActiveBar2LibraryCtl.DockingAreaTypes = 3&, Optional ByVal paramGrabHandleStyle As ActiveBar2LibraryCtl.GrabHandleStyles = 7&, Optional ByVal paramDockingOffset As Long = 0&) As ActiveBar2LibraryCtl.IBand
'Dim b As ActiveBar2LibraryCtl.Band
'Dim t As ActiveBar2LibraryCtl.tool
'Dim sBandName As String
'
'On Error GoTo eh_IWillDockToActiveBar_DockYourselfTo
'
'
'    sBandName = DOCKABLEBANDPREFIXNAME + Me.Name
'
'    ' The Dockable Form band for this form does not exist, so create one.
'    Set b = ActiveBar.Bands.Add(sBandName)
'        b.Caption = Me.Caption
'
'        b.DockingArea = paramDockingarea
'        b.DockLine = 0
'        b.DockingOffset = paramDockingOffset
'
'        b.GrabHandleStyle = paramGrabHandleStyle
'
'        b.AutoSizeForms = True
'        b.Type = ddBTNormal
'        b.DisplayMoreToolsButton = False
'
'        ABAddFlag ddBFSizer, b
'
'        b.Visible = parmIsVisible
'
'    ' Add a DockableForm tool to dock this window to.
'    Set t = b.Tools.Add(Me.hWnd, DOCKABLETOOLPREFIXNAME + Me.Name)
'        t.ControlType = ddTTForm
'        t.Caption = Me.Caption
'        Set t.Custom = Me
'
'
'ex_IWillDockToActiveBar_DockYourselfTo:
'
'Exit Function
'
'eh_IWillDockToActiveBar_DockYourselfTo:
'
'    MsgBox "There was an error while docking form [" + Me.Name + "]."
'
'    Resume ex_IWillDockToActiveBar_DockYourselfTo
'End Function




