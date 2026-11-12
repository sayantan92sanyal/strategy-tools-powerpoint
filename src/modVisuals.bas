Attribute VB_Name = "modVisuals"
' Strategy Tools - one-click visuals: Harvey balls and rising-bar meters.
' Each visual is a group tagged ST_TYPE (HARVEY / METER) and ST_LEVEL, so
' selecting an existing one and picking a new level updates it in place.
Option Explicit

Private Const VISUAL_SIZE As Single = 28.8   ' 0.4 inch
Private Const BAR_WIDTH As Single = 8
Private Const BAR_GAP As Single = 2.4

' ---------------- Harvey balls ----------------

Public Sub ST_Harvey0():   HarveyBall 0:   End Sub
Public Sub ST_Harvey25():  HarveyBall 25:  End Sub
Public Sub ST_Harvey50():  HarveyBall 50:  End Sub
Public Sub ST_Harvey75():  HarveyBall 75:  End Sub
Public Sub ST_Harvey100(): HarveyBall 100: End Sub

Public Sub HarveyBall(ByVal pct As Long)
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim targets As Collection
    Set targets = SelectedVisuals("HARVEY")
    If targets.Count > 0 Then
        Dim g As Variant
        For Each g In targets
            StyleHarvey g, pct
        Next g
    Else
        Dim sld As Object, base As Object, pie As Object, grp As Object
        Set sld = CurrentSlide()
        Dim x As Single, y As Single
        x = (ActivePresentation.PageSetup.SlideWidth - VISUAL_SIZE) / 2
        y = (ActivePresentation.PageSetup.SlideHeight - VISUAL_SIZE) / 2
        Set base = sld.Shapes.AddShape(ST_SHAPE_OVAL, x, y, VISUAL_SIZE, VISUAL_SIZE)
        Set pie = sld.Shapes.AddShape(ST_SHAPE_PIE, x, y, VISUAL_SIZE, VISUAL_SIZE)
        Set grp = GroupShapes(sld, Array(base, pie), Array("ST_Base", "ST_Fill"))
        grp.Tags.Add "ST_TYPE", "HARVEY"
        StyleHarvey grp, pct
        grp.Select
    End If
    Exit Sub
fail:
    ReportError "Harvey ball"
End Sub

Private Sub StyleHarvey(ByVal grp As Object, ByVal pct As Long)
    Dim base As Object, pie As Object, item As Object
    For Each item In grp.GroupItems
        If item.Name = "ST_Base" Then Set base = item
        If item.Name = "ST_Fill" Then Set pie = item
    Next item
    NoOutline base
    NoOutline pie
    base.Fill.Solid
    pie.Fill.Solid
    pie.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
    If pct >= 100 Then
        base.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
        pie.Visible = ST_FALSE
    Else
        base.Fill.ForeColor.RGB = GreyRGB()
        If pct <= 0 Then
            pie.Visible = ST_FALSE
        Else
            pie.Visible = ST_TRUE
            ' Angles run clockwise from 3 o'clock; 270 = 12 o'clock.
            pie.Adjustments(1) = 270
            pie.Adjustments(2) = (270 + 360 * pct / 100) Mod 360
        End If
    End If
    grp.Tags.Add "ST_LEVEL", CStr(pct)
    grp.Name = "Harvey Ball " & pct & "%"
End Sub

' ---------------- Meter ----------------

Public Sub ST_MeterLow():    Meter 1: End Sub
Public Sub ST_MeterMedium(): Meter 2: End Sub
Public Sub ST_MeterHigh():   Meter 3: End Sub

Public Sub Meter(ByVal level As Long)
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim targets As Collection
    Set targets = SelectedVisuals("METER")
    If targets.Count > 0 Then
        Dim g As Variant
        For Each g In targets
            StyleMeter g, level
        Next g
    Else
        Dim sld As Object, grp As Object, i As Long
        Dim bars(1 To 3) As Object, h As Single, totalW As Single
        Set sld = CurrentSlide()
        totalW = 3 * BAR_WIDTH + 2 * BAR_GAP
        Dim x As Single, bottom As Single
        x = (ActivePresentation.PageSetup.SlideWidth - totalW) / 2
        bottom = (ActivePresentation.PageSetup.SlideHeight + VISUAL_SIZE) / 2
        For i = 1 To 3
            h = VISUAL_SIZE * (i + 1) / 4          ' 50%, 75%, 100% of full height
            Set bars(i) = sld.Shapes.AddShape(ST_SHAPE_RECTANGLE, _
                x + (i - 1) * (BAR_WIDTH + BAR_GAP), bottom - h, BAR_WIDTH, h)
        Next i
        Set grp = GroupShapes(sld, Array(bars(1), bars(2), bars(3)), Array("ST_Bar1", "ST_Bar2", "ST_Bar3"))
        grp.Tags.Add "ST_TYPE", "METER"
        StyleMeter grp, level
        grp.Select
    End If
    Exit Sub
fail:
    ReportError "Meter"
End Sub

Private Sub StyleMeter(ByVal grp As Object, ByVal level As Long)
    Dim item As Object, n As Long
    For Each item In grp.GroupItems
        If Left$(item.Name, 6) = "ST_Bar" Then
            n = CLng(Mid$(item.Name, 7))
            NoOutline item
            item.Fill.Solid
            If n <= level Then
                item.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
            Else
                item.Fill.ForeColor.RGB = GreyRGB()
            End If
        End If
    Next item
    grp.Tags.Add "ST_LEVEL", CStr(level)
    grp.Name = "Meter " & Choose(level, "Low", "Medium", "High")
End Sub

' ---------------- helpers ----------------

' Group freshly added shapes, then give the parts their fixed names.
' Shapes are grouped by unique temporary names because index/z-order numbers
' shift once the slide already contains groups.
Private Function GroupShapes(ByVal sld As Object, ByVal shapes As Variant, ByVal finalNames As Variant) As Object
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

' Selected top-level groups whose ST_TYPE tag matches kind.
Private Function SelectedVisuals(ByVal kind As String) As Collection
    Dim result As New Collection
    Dim sr As Object, i As Long, s As Object
    Set sr = SelectedShapes()
    If Not sr Is Nothing Then
        For i = 1 To sr.Count
            Set s = TopLevel(sr(i))
            If TagOf(s, "ST_TYPE") = kind Then result.Add s
        Next i
    End If
    Set SelectedVisuals = result
End Function

Private Sub NoOutline(ByVal shp As Object)
    shp.Line.Visible = ST_FALSE
End Sub
