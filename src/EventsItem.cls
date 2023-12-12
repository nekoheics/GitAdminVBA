VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MenuItem As VBIDE.CommandBarEvents
Attribute MenuItem.VB_VarHelpID = -1

Private Sub Class_Terminate()
    Set MenuItem = Nothing
End Sub

Public Sub Initialize(ByVal ctl As Office.CommandBarControl)
    Set MenuItem = Application.VBE.Events.CommandBarEvents(ctl)
End Sub

Private Sub MenuItem_Click(ByVal sender As Object, handled As Boolean, CancelDefault As Boolean)
    Application.Run sender.OnAction
End Sub
