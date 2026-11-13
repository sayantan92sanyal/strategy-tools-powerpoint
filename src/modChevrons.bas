Attribute VB_Name = "modChevrons"
' Strategy Tools - phase chevrons with a detail text box under each.
' Laid out across the title's width (or the slide's, with margins), just below
' the title, as one group tagged ST_TYPE = CHEVRONS.
Option Explicit

Private Const SHAPE_PENTAGON As Long = 51    ' msoShapePentagon (flat left, pointed right)
Private Const SHAPE_CHEVRON As Long = 52     ' msoShapeChevron
Private Const POINT_RATIO As Single = 0.35   ' arrow-point depth as a share of chevron height
Private Const GAP As Single = 4
Private Const MIN_COUNT As Long = 2
Private Const MAX_COUNT As Long = 8

' Ribbon button: ask how many, then draw.
Public Sub ST_Chevrons()
    On Error GoTo fail
    If Not HasEditableSlide() Then
        Msg "Open a slide in Normal view first."
        Exit Sub
    End If
    Dim answer As String
    answer = InputBox("How many chevrons? (" & MIN_COUNT & " to " & MAX_COUNT & ")", ST_APP_TITLE, "4")
    If Len(Trim$(answer)) = 0 Then Exit Sub          ' Cancel
    If Not IsNumeric(answer) Then
        Msg "Please enter a number from " & MIN_COUNT & " to " & MAX_COUNT & "."
        Exit Sub
    End If
    AddChevrons CLng(answer)
    Exit Sub
fail:
    ReportError "Chevrons"
End Sub

Public Sub AddChevrons(ByVal n As Long)
    On Error GoTo fail
    If n < MIN_COUNT Or n > MAX_COUNT Then
        Msg "Please enter a number from " & MIN_COUNT & " to " & MAX_COUNT & "."
        Exit Sub
    End If
    Dim sld As Object, slideW As Single, slideH As Single
    Set sld = CurrentSlide()
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    ' Area: line up with the title if there is one, else use 6% margins.
    Dim areaLeft As Single, areaWidth As Single, areaTop As Single
    areaLeft = slideW * 0.06: areaWidth = slideW * 0.88: areaTop = slideH * 0.2
    On Error Resume Next
    If sld.Shapes.HasTitle = ST_TRUE Then
        With sld.Shapes.Title
            If .Width > slideW * 0.5 Then areaLeft = .Left: areaWidth = .Width
            If .Top + .Height < slideH * 0.35 Then areaTop = .Top + .Height + 12
        End With
    End If
    On Error GoTo fail

    Dim h As Single, depth As Single, w As Single, stepX As Single
    h = slideH * 0.09
    depth = h * POINT_RATIO
    ' n chevrons overlap by (depth - GAP) each so their notches nest.
    w = (areaWidth + (n - 1) * (depth - GAP)) / n
    stepX = w - depth + GAP

    Dim parts() As Variant, finals() As Variant, i As Long, shp As Object, box As Object
    ReDim parts(1 To 2 * n)
    ReDim finals(1 To 2 * n)
    For i = 1 To n
        Set shp = sld.Shapes.AddShape(IIf(i = 1, SHAPE_PENTAGON, SHAPE_CHEVRON), _
                                      areaLeft + (i - 1) * stepX, areaTop, w, h)
        shp.Adjustments(1) = POINT_RATIO
        shp.Fill.Solid
        shp.Fill.ForeColor.ObjectThemeColor = ST_THEME_ACCENT1
        NoOutline shp
        With shp.TextFrame
            .WordWrap = ST_TRUE
            .MarginLeft = depth: .MarginRight = depth / 2
            With .TextRange
                .Text = "Phase " & i
                .ParagraphFormat.Alignment = ppAlignCenter
                .Font.Bold = ST_TRUE
                .Font.Size = 14
                .Font.Color.RGB = RGB(255, 255, 255)
            End With
        End With
        Set parts(i) = shp
        finals(i) = "ST_Chevron" & i

        ' Detail box: one column per chevron, same pitch, below the row.
        Set box = sld.Shapes.AddTextbox(ST_TEXT_HORIZONTAL, areaLeft + (i - 1) * stepX + IIf(i = 1, 0, depth / 2), _
                                        areaTop + h + 10, stepX - GAP - IIf(i = 1, 0, depth / 2) + depth / 2, slideH * 0.4)
        With box.TextFrame
            .WordWrap = ST_TRUE
            .AutoSize = ppAutoSizeNone
            .TextRange.Text = "Detail" & vbCr & "Detail"
            .TextRange.Font.Size = 12
            With .TextRange.ParagraphFormat.Bullet
                .Visible = ST_TRUE
                .Type = ppBulletUnnumbered
                .Character = 8226          ' round bullet
            End With
            .TextRange.ParagraphFormat.Alignment = ppAlignLeft
        End With
        Set parts(n + i) = box
        finals(n + i) = "ST_Detail" & i
    Next i

    Dim grp As Object
    Set grp = GroupShapes(sld, parts, finals)
    grp.Tags.Add "ST_TYPE", "CHEVRONS"
    grp.Tags.Add "ST_LEVEL", CStr(n)
    grp.Name = "Chevrons (" & n & ")"
    grp.Select
    Exit Sub
fail:
    ReportError "Chevrons"
End Sub
