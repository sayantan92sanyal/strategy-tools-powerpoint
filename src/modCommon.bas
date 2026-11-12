Attribute VB_Name = "modCommon"
' Strategy Tools - shared helpers used by every feature module.
Option Explicit

' Office constants (the Office type library is not referenced, so we define
' the values we need here).
Public Const ST_TRUE As Long = -1
Public Const ST_FALSE As Long = 0
Public Const ST_SHAPE_RECTANGLE As Long = 1
Public Const ST_SHAPE_OVAL As Long = 9
Public Const ST_SHAPE_PIE As Long = 142
Public Const ST_THEME_ACCENT1 As Long = 5
Public Const ST_TYPE_GROUP As Long = 6
Public Const ST_TYPE_PLACEHOLDER As Long = 14
Public Const ST_TEXT_HORIZONTAL As Long = 1

Public Const ST_APP_TITLE As String = "Strategy Tools"

' When True (set by tests), messages are recorded instead of shown.
Public gQuiet As Boolean
Public gLastMessage As String

' Light grey used for the "empty" part of Harvey balls and meters.
Public Function GreyRGB() As Long
    GreyRGB = RGB(227, 230, 234)
End Function

Public Sub Msg(ByVal text As String)
    gLastMessage = text
    If Not gQuiet Then MsgBox text, vbInformation, ST_APP_TITLE
End Sub

' True when there is a normal editing window with a slide showing.
Public Function HasEditableSlide() As Boolean
    On Error GoTo nope
    If Presentations.Count = 0 Then GoTo nope
    If ActiveWindow Is Nothing Then GoTo nope
    Dim s As Object
    Set s = ActiveWindow.View.Slide
    HasEditableSlide = Not (s Is Nothing)
    Exit Function
nope:
    HasEditableSlide = False
End Function

Public Function CurrentSlide() As Object
    Set CurrentSlide = ActiveWindow.View.Slide
End Function

' Returns the selected shapes as a ShapeRange in selection order, or Nothing.
Public Function SelectedShapes() As Object
    On Error GoTo nope
    Set SelectedShapes = Nothing
    If Presentations.Count = 0 Then Exit Function
    Dim t As Long
    t = ActiveWindow.Selection.Type
    If t = ppSelectionShapes Or t = ppSelectionText Then
        Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    End If
    Exit Function
nope:
    Set SelectedShapes = Nothing
End Function

Public Function SelectionCount() As Long
    Dim sr As Object
    Set sr = SelectedShapes()
    If sr Is Nothing Then SelectionCount = 0 Else SelectionCount = sr.Count
End Function

' If a shape inside a group is selected, return the whole group.
Public Function TopLevel(ByVal shp As Object) As Object
    On Error GoTo plain
    If shp.Child = ST_TRUE Then
        Set TopLevel = shp.ParentGroup
        Exit Function
    End If
plain:
    Set TopLevel = shp
End Function

Public Function TagOf(ByVal shp As Object, ByVal tagName As String) As String
    On Error Resume Next
    TagOf = shp.Tags(tagName)
End Function

Public Sub ReportError(ByVal where As String)
    Msg "Sorry, " & where & " didn't work: " & Err.Description
End Sub
