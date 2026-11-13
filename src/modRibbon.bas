Attribute VB_Name = "modRibbon"
' Strategy Tools - ribbon button callbacks. Each one just calls the feature
' macro, so the same feature can also be run from a keyboard shortcut.
' To add a feature: write a Public Sub ST_Something() in a feature module,
' add a Rb_ callback here, and add a <button> to src/customUI14.xml.
Option Explicit

Public Sub Rb_Swap(control As Object):          ST_Swap:            End Sub
Public Sub Rb_SameSize(control As Object):      ST_SameSize:        End Sub
Public Sub Rb_SameWidth(control As Object):     ST_SameWidth:       End Sub
Public Sub Rb_SameHeight(control As Object):    ST_SameHeight:      End Sub
Public Sub Rb_TitleHarmonizer(control As Object): ST_TitleHarmonizer: End Sub

' Harvey ball menu: the button's tag holds the percentage.
Public Sub Rb_Harvey(control As Object)
    HarveyBall CLng(control.Tag)
End Sub

' Meter menu: the button's tag holds the level (1 = Low, 2 = Medium, 3 = High).
Public Sub Rb_Meter(control As Object)
    Meter CLng(control.Tag)
End Sub

' Status menu: the button's tag holds COMPLETE / INPROGRESS / DELAYED / NOTSTARTED.
Public Sub Rb_Status(control As Object)
    StatusIcon CStr(control.Tag)
End Sub

Public Sub Rb_Chevrons(control As Object):      ST_Chevrons:        End Sub
Public Sub Rb_TrackerNumber(control As Object): ST_TrackerNumber:   End Sub
Public Sub Rb_TrackerLetter(control As Object): ST_TrackerLetter:   End Sub
Public Sub Rb_TrackerCustom(control As Object): ST_TrackerCustom:   End Sub

Public Sub Rb_TextBox(control As Object):       ST_AddTextBox:      End Sub
Public Sub Rb_Shortcuts(control As Object):     ST_ShowShortcuts:   End Sub
