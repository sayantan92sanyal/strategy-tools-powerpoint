Attribute VB_Name = "modStatus"
' Strategy Tools - status icons: Complete, In progress, Delayed, Not started.
' Fixed traffic-light colours. Each icon is a group tagged ST_TYPE = STATUS;
' selecting existing icons and picking a status redraws them in place.
Option Explicit

Private Const ICON_SIZE As Single = 28.8     ' 0.4 inch

Public Sub ST_StatusComplete():   StatusIcon "COMPLETE":   End Sub
Public Sub ST_StatusInProgress(): StatusIcon "INPROGRESS": End Sub
Public Sub ST_StatusDelayed():    StatusIcon "DELAYED":    End Sub
Public Sub ST_StatusNotStarted(): StatusIcon "NOTSTARTED": End Sub

Public Sub StatusIcon(ByVal status As String)
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim sld As Object, targets As Collection, g As Variant
    Dim x As Single, y As Single, w As Single, made As Object
    Set sld = CurrentSlide()
    Set targets = SelectedVisuals("STATUS")
    If targets.Count > 0 Then
        Dim names() As Variant, n As Long
        ReDim names(1 To targets.Count)
        For Each g In targets
            x = g.Left: y = g.Top: w = g.Width
            g.Delete
            Set made = DrawStatus(sld, status, x, y, w)
            n = n + 1
            names(n) = made.Name
        Next g
        sld.Shapes.Range(names).Select
    Else
        x = (ActivePresentation.PageSetup.SlideWidth - ICON_SIZE) / 2
        y = (ActivePresentation.PageSetup.SlideHeight - ICON_SIZE) / 2
        Set made = DrawStatus(sld, status, x, y, ICON_SIZE)
        made.Select
    End If
    Exit Sub
fail:
    ReportError "Status icon"
End Sub

' Draws one icon of size s at (x, y) and returns the group.
Private Function DrawStatus(ByVal sld As Object, ByVal status As String, _
                            ByVal x As Single, ByVal y As Single, ByVal s As Single) As Object
    Dim disc As Object, marks As Collection, shp As Object, colour As Long, label As String
    Set marks = New Collection
    Select Case status
        Case "COMPLETE":   colour = RGB(46, 158, 91):  label = "Complete"
        Case "INPROGRESS": colour = RGB(240, 162, 2):  label = "In progress"
        Case "DELAYED":    colour = RGB(214, 69, 69):  label = "Delayed"
        Case Else:         colour = RGB(163, 170, 181): label = "Not started": status = "NOTSTARTED"
    End Select

    Set disc = sld.Shapes.AddShape(ST_SHAPE_OVAL, x, y, s, s)
    disc.Fill.Solid
    disc.Fill.ForeColor.RGB = colour
    NoOutline disc

    ' Mark coordinates are on a 56-unit grid, scaled to the icon size.
    Select Case status
        Case "COMPLETE"
            marks.Add Tick(sld, x, y, s)
        Case "INPROGRESS"
            AddDottedTick sld, x, y, s, marks
        Case "DELAYED"
            marks.Add Stroke(sld, x, y, s, 19, 19, 37, 37)
            marks.Add Stroke(sld, x, y, s, 37, 19, 19, 37)
        Case Else
            marks.Add Stroke(sld, x, y, s, 18, 28, 38, 28)
    End Select

    Dim parts() As Variant, finals() As Variant, i As Long
    ReDim parts(0 To marks.Count)
    ReDim finals(0 To marks.Count)
    Set parts(0) = disc: finals(0) = "ST_Circle"
    For i = 1 To marks.Count
        Set parts(i) = marks(i)
        finals(i) = "ST_Mark" & i
    Next i
    Dim grp As Object
    Set grp = GroupShapes(sld, parts, finals)
    grp.Tags.Add "ST_TYPE", "STATUS"
    grp.Tags.Add "ST_LEVEL", status
    grp.Name = "Status " & label
    Set DrawStatus = grp
End Function

Private Function Tick(ByVal sld As Object, ByVal x As Single, ByVal y As Single, ByVal s As Single) As Object
    Dim pts(1 To 3, 1 To 2) As Single, k As Single
    k = s / 56
    pts(1, 1) = x + 16 * k: pts(1, 2) = y + 29 * k
    pts(2, 1) = x + 25 * k: pts(2, 2) = y + 38 * k
    pts(3, 1) = x + 41 * k: pts(3, 2) = y + 20 * k
    Dim shp As Object
    Set shp = sld.Shapes.AddPolyline(pts)
    shp.Fill.Visible = ST_FALSE
    WhiteLine shp, s
    Set Tick = shp
End Function

' The tick path drawn as evenly spaced white dots (same 56-unit grid as Tick).
Private Sub AddDottedTick(ByVal sld As Object, ByVal x As Single, ByVal y As Single, _
                          ByVal s As Single, ByVal marks As Collection)
    Dim px As Variant, py As Variant, seg As Long, k As Single, d As Single
    Dim segLen As Single, nDots As Long, i As Long, f As Single, cx As Single, cy As Single
    Dim dot As Object
    px = Array(16, 25, 41)
    py = Array(29, 38, 20)
    k = s / 56
    d = 5 * k                                   ' dot diameter = tick line thickness
    For seg = 0 To 1
        segLen = Sqr((px(seg + 1) - px(seg)) ^ 2 + (py(seg + 1) - py(seg)) ^ 2)
        nDots = Int(segLen / 8.5 + 0.5)          ' about one dot every 8.5 units
        For i = IIf(seg = 0, 0, 1) To nDots
            f = i / nDots
            cx = x + (px(seg) + (px(seg + 1) - px(seg)) * f) * k
            cy = y + (py(seg) + (py(seg + 1) - py(seg)) * f) * k
            Set dot = sld.Shapes.AddShape(ST_SHAPE_OVAL, cx - d / 2, cy - d / 2, d, d)
            dot.Fill.Solid
            dot.Fill.ForeColor.RGB = RGB(255, 255, 255)
            NoOutline dot
            marks.Add dot
        Next i
    Next seg
End Sub

Private Function Stroke(ByVal sld As Object, ByVal x As Single, ByVal y As Single, ByVal s As Single, _
                        ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Object
    Dim k As Single, shp As Object
    k = s / 56
    Set shp = sld.Shapes.AddLine(x + x1 * k, y + y1 * k, x + x2 * k, y + y2 * k)
    WhiteLine shp, s
    Set Stroke = shp
End Function

Private Sub WhiteLine(ByVal shp As Object, ByVal s As Single)
    With shp.Line
        .Visible = ST_TRUE
        .ForeColor.RGB = RGB(255, 255, 255)
        .Weight = s * 0.09
    End With
End Sub
