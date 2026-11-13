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

' Group freshly added shapes, then give the parts their fixed names.
' Shapes are grouped by unique temporary names because index/z-order numbers
' shift once the slide already contains groups.
Public Function GroupShapes(ByVal sld As Object, ByVal shapes As Variant, ByVal finalNames As Variant) As Object
    Dim i As Long, stamp As String, names() As Variant
    stamp = "ST_tmp_" & Format(Now, "hhmmss") & "_" & CLng(Rnd * 1000000) & "_"
    ReDim names(LBound(shapes) To UBound(shapes))
    For i = LBound(shapes) To UBound(shapes)
        names(i) = stamp & i
        shapes(i).Name = names(i)
    Next i
    Dim grp As Object, item As Object
    Set grp = sld.Shapes.Range(names).Group
    For Each item In grp.GroupItems
        For i = LBound(names) To UBound(names)
            If item.Name = names(i) Then item.Name = finalNames(i)
        Next i
    Next item
    Set GroupShapes = grp
End Function

' Selected top-level shapes whose ST_TYPE tag matches kind (each once).
Public Function SelectedVisuals(ByVal kind As String) As Collection
    Dim result As New Collection
    Dim sr As Object, i As Long, s As Object
    Set sr = SelectedShapes()
    If Not sr Is Nothing Then
        On Error Resume Next   ' Collection key de-duplicates groups
        For i = 1 To sr.Count
            Set s = TopLevel(sr(i))
            If TagOf(s, "ST_TYPE") = kind Then result.Add s, s.Name & "|" & s.Id
        Next i
        On Error GoTo 0
    End If
    Set SelectedVisuals = result
End Function

' Slides to act on: the slides selected in the thumbnail pane, or else the
' slide being edited. Empty when no slide is available.
Public Function TargetSlides() As Collection
    Dim c As New Collection, s As Object, rng As Object
    Set rng = SelectedSlideRange()
    If Not rng Is Nothing Then
        For Each s In rng
            c.Add s
        Next s
    End If
    If c.Count = 0 Then
        If HasEditableSlide() Then c.Add CurrentSlide()
    End If
    Set TargetSlides = c
End Function

Private Function SelectedSlideRange() As Object
    On Error GoTo none
    If ActiveWindow.Selection.Type = ppSelectionSlides Then Set SelectedSlideRange = ActiveWindow.Selection.SlideRange
none:
End Function

' Yes/No question. Tests (gQuiet) always answer Yes.
Public Function Confirm(ByVal text As String) As Boolean
    gLastMessage = text
    If gQuiet Then
        Confirm = True
    Else
        Confirm = (MsgBox(text, vbYesNo + vbQuestion, ST_APP_TITLE) = vbYes)
    End If
End Function

' Shapes on a slide carrying ST_TYPE = kind (top level only).
Public Function TaggedShapes(ByVal sld As Object, ByVal kind As String) As Collection
    Dim c As New Collection, shp As Object
    For Each shp In sld.Shapes
        If TagOf(shp, "ST_TYPE") = kind Then c.Add shp
    Next shp
    Set TaggedShapes = c
End Function

' Left edge used by consulting elements: the master title's left, else 4%.
Public Function ContentLeft() As Single
    Dim shp As Object
    Dim i As Long, n As Long
    On Error GoTo done
    ContentLeft = ActivePresentation.PageSetup.SlideWidth * 0.04
    n = ActivePresentation.SlideMaster.Shapes.Count
    For i = 1 To n
        Set shp = ActivePresentation.SlideMaster.Shapes(i)
        If IsPlaceholderOf(shp, ppPlaceholderTitle) Then ContentLeft = shp.Left: Exit For
    Next i
done:
End Function

' Test-only trace: appends a line to st_trace.log in PowerPoint's temp folder
' when running under tests (gQuiet). Does nothing for normal use.
Public Sub Trace(ByVal text As String)
    If Not gQuiet Then Exit Sub
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ("TMPDIR") & "st_trace.log" For Append As #f
    Print #f, Format(Now, "hh:nn:ss") & " " & text
    Close #f
End Sub

Public Function IsPlaceholderOf(ByVal shp As Object, ByVal kind As Long) As Boolean
    On Error GoTo nope
    If shp.Type <> ST_TYPE_PLACEHOLDER Then Exit Function
    IsPlaceholderOf = (shp.PlaceholderFormat.Type = kind)
nope:
End Function
